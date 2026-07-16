using System.Buffers.Text;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
{
    public sealed partial class XlsxReader
    {
        // Forward-only, low-memory worksheet scanner. Streams the sheet through a refillable pooled
        // buffer; a single <c>...</c> element is guaranteed contiguous (the buffer grows if needed).
        [SuppressMessage("Design", "CA1034:Nested types should not be visible",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public sealed class Enumerator : IExcelRowEnumerator
        {
            // Borrowed: the reader outlives the enumerator and owns its own disposal — do not dispose here.
            [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed", Justification = "Borrowed, not owned.")]
            [SuppressMessage("SharpSource", "SS066:Disposable field is not disposed", Justification = "Borrowed, not owned.")]
            private readonly XlsxReader _reader;
            private readonly CancellationToken _ct; // honored only by the async path
            // Owned: opened by Get(Async)Enumerator for this enumerator alone; disposed in Dispose(Async).
            [SuppressMessage("SharpSource", "SS066:Disposable field is not disposed", Justification = "Disposed in Dispose().")]
            private Stream? _sheet;
            private readonly BufferedStreamCursor _io;
            private byte[] _buf => _io.Buf;
            private int _pos { get => _io.Pos; set => _io.Pos = value; }
            private int _len => _io.Len;
            private bool _eof => _io.Eof;

            private readonly CellAccumulator _acc; // per-row decoded values + cell descriptors
            private int _nextCol;

            // Non-null only when this sheet's elements carry a namespace prefix (e.g. <x:row>); holds the
            // prefixed forms of every token the scanner matches. Detected once, lazily, on the first
            // MoveNext(Async). Null for the default-namespace case, which keeps the literal fast paths.
            private NsTokens? _ns;
            private bool _nsChecked;

            private ReadOnlySpan<byte> VOpen => _ns is null ? "<v>"u8 : _ns.VOpen;
            private ReadOnlySpan<byte> VClose => _ns is null ? "</v>"u8 : _ns.VClose;
            private ReadOnlySpan<byte> CClose => _ns is null ? "</c>"u8 : _ns.CClose;

            internal Enumerator(XlsxReader reader, Stream sheet, long entryLength = 0, CancellationToken ct = default)
            {
                _reader = reader;
                _sheet = sheet;
                _ct = ct;
                _io = new BufferedStreamCursor(reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes),
                    WorkbookLookups.InitialBufferCapacity(entryLength));
                _acc = new CellAccumulator(reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes));
            }

            public Row Current =>
                new(_acc.CellSpan, _acc.ValueSpan, _reader.SharedSpan, _reader.SharedStringCache);

            // Top-level scanning (finding the next '<row>'/'</sheetData>'/markup between rows) differs
            // between sync and async only in whether the buffer refill (Fill / FillAsync) awaits — so
            // that span work is factored into sync helpers (ClassifyHead, BeginRow) that never hold a
            // span across an await. Once inside a row, ParseRow/ParseCellSpan/EmitCell are the *same*
            // method for both: EnsureRowBuffered(Async) guarantees the whole row is buffered first, so
            // there is no refill left to do and no sync/async split needed at all (see T3.1 note below).

            public bool MoveNext()
            {
                if (!_nsChecked)
                {
                    DetectNamespace();
                }
                while (true)
                {
                    // Fast path: in compact output _pos already sits on the next '<', so skip the scan.
                    int lt = _pos < _len && _buf[_pos] == (byte)'<' ? _pos : IndexOf((byte)'<');
                    if (lt < 0)
                    {
                        return false;
                    }
                    _pos = lt;
                    Ensure(_ns is null ? 12 : _ns.HeadEnsure);
                    switch (ClassifyHead())
                    {
                        case HeadKind.End:
                            return false;
                        case HeadKind.Row:
                            if (!BeginRow())
                            {
                                ParseRow(EnsureRowBuffered());
                            }
                            return true;
                        default:
                            if (!SkipMarkup())
                            {
                                return false;
                            }
                            break;
                    }
                }
            }

            // Non-async fast path: IndexOfAsync/EnsureAsync/BeginRowAsync/EnsureRowBufferedAsync/
            // SkipMarkupAsync are each already "check synchronously, only await on a genuine buffer
            // miss" — but the original method being itself `async` meant every row paid for a state
            // machine regardless. This mirrors the same steps but returns a completed ValueTask when
            // every step resolves synchronously (~99.9% of rows), only falling to an awaiting
            // continuation at the exact step that needs a refill.
            [SuppressMessage("SharpSource", "SS034:Use await to get the result of a Task",
                Justification = "Every .Result access is guarded by IsCompletedSuccessfully immediately above it — never blocks.")]
            [SuppressMessage("VisualStudio.Threading", "VSTHRD103:Result synchronously blocks",
                Justification = "Every .Result access is guarded by IsCompletedSuccessfully immediately above it — never blocks.")]
            public ValueTask<bool> MoveNextAsync()
            {
                if (!_nsChecked)
                {
                    return DetectNamespaceThenMoveNextAsync();
                }
                while (true)
                {
                    int lt;
                    if (_pos < _len && _buf[_pos] == (byte)'<')
                    {
                        lt = _pos;
                    }
                    else
                    {
                        ValueTask<int> ltTask = IndexOfAsync((byte)'<');
                        if (!ltTask.IsCompletedSuccessfully)
                        {
                            return AwaitThenRestartAsync(ltTask);
                        }
                        lt = ltTask.Result;
                    }
                    if (lt < 0)
                    {
                        return new ValueTask<bool>(false);
                    }
                    _pos = lt;

                    ValueTask ensureTask = EnsureAsync(_ns is null ? 12 : _ns.HeadEnsure);
                    if (!ensureTask.IsCompletedSuccessfully)
                    {
                        return AwaitThenRestartAsync(ensureTask);
                    }

                    switch (ClassifyHead())
                    {
                        case HeadKind.End:
                            return new ValueTask<bool>(false);
                        case HeadKind.Row:
                            {
                                ValueTask<bool> beginTask = BeginRowAsync();
                                if (!beginTask.IsCompletedSuccessfully)
                                {
                                    return AwaitThenRestartAsync(beginTask);
                                }
                                if (!beginTask.Result)
                                {
                                    ValueTask<int> rowBufTask = EnsureRowBufferedAsync();
                                    if (!rowBufTask.IsCompletedSuccessfully)
                                    {
                                        return FinishRowAfterAsync(rowBufTask);
                                    }
                                    ParseRow(rowBufTask.Result);
                                }
                                return new ValueTask<bool>(true);
                            }
                        default:
                            {
                                ValueTask<bool> skipTask = SkipMarkupAsync();
                                if (!skipTask.IsCompletedSuccessfully)
                                {
                                    return AwaitThenRestartAsync(skipTask);
                                }
                                if (!skipTask.Result)
                                {
                                    return new ValueTask<bool>(false);
                                }
                                break;
                            }
                    }
                }
            }

            // Safe for every pending step above except the row-buffered check below: none of them
            // commit a position change until they resolve (BeginRowAsync only advances _pos once it
            // actually finds '>'; SkipMarkupAsync only advances _pos at its return statement). So once
            // the fill completes, simply re-entering MoveNextAsync redoes the (now-buffered, cheap) work.
            private async ValueTask<bool> AwaitThenRestartAsync(ValueTask pending)
            {
                await pending.ConfigureAwait(false);
                return await MoveNextAsync().ConfigureAwait(false);
            }

            private async ValueTask<bool> AwaitThenRestartAsync<T>(ValueTask<T> pending)
            {
                await pending.ConfigureAwait(false);
                return await MoveNextAsync().ConfigureAwait(false);
            }

            // BeginRowAsync already succeeded here (row is not self-closed, _pos is now past <row ...>),
            // so unlike AwaitThenRestartAsync this must not re-enter MoveNextAsync from the top — that
            // would misread the row's first cell as a new top-level element. Finishes buffering the row,
            // then parses it.
            private async ValueTask<bool> FinishRowAfterAsync(ValueTask<int> pendingRowBuffered)
            {
                int rowEnd = await pendingRowBuffered.ConfigureAwait(false);
                ParseRow(rowEnd);
                return true;
            }

            // Detects the sheet's element-name prefix (e.g. "x:" in <x:worksheet>) exactly once, from the
            // root element at the very start of the stream. Prefixed worksheets are rare (Excel emits the
            // default namespace), so this stays out of the per-row/per-cell hot path — it runs once, then
            // _ns is null (fast literal matching) or holds the prefixed tokens for the whole enumeration.
            private void DetectNamespace()
            {
                _nsChecked = true;
                Ensure(256); // root element + its xmlns declarations sit at the head of the part
                DetectNamespaceFromBuffer();
            }

            private async ValueTask<bool> DetectNamespaceThenMoveNextAsync()
            {
                _nsChecked = true;
                await EnsureAsync(256).ConfigureAwait(false);
                DetectNamespaceFromBuffer();
                return await MoveNextAsync().ConfigureAwait(false);
            }

            private void DetectNamespaceFromBuffer()
            {
                ReadOnlySpan<byte> prefix = XlsxXml.DetectElementPrefix(_buf.AsSpan(_pos, _len - _pos));
                if (!prefix.IsEmpty)
                {
                    _ns = new NsTokens(prefix);
                }
            }

            // Consumes the <row ...> open tag and resets per-row state. Call only after ClassifyHead()==Row.
            private bool BeginRow()
            {
                int gt = IndexOf((byte)'>'); // open tag already fully buffered by the Ensure(12) above
                if (gt < 0)
                {
                    return MissingRowOpenTag();
                }
                return BeginRowAt(gt);
            }

            // Whole-row span parser: `rowEnd` (the '<' that starts "</row") is supplied by
            // EnsureRowBuffered/EnsureRowBufferedAsync, which already grew the buffer (via Fill/FillAsync)
            // until the entire row — cell data plus "</row...>"'s closing '>' — is present. Everything
            // below is then pure ReadOnlySpan<byte> work over a local cursor: zero Ensure/Fill calls, zero
            // _io property indirection, and no risk of PrepareBuffer compacting mid-row (which is what
            // made the old per-cell fast path need careful `_pos`-relative rewinds). `_io.Pos` is written
            // back exactly once, after the whole row is consumed. Called identically by MoveNext (sync)
            // and MoveNextAsync (after awaiting EnsureRowBufferedAsync) — one parser, no async twin.
            private void ParseRow(int rowEnd)
            {
                byte[] buf = _buf;
                int len = _len;
                int p = _pos;
                while (true)
                {
                    int lt = p < len && buf[p] == (byte)'<' ? p : IndexOfBounded(buf, len, p, (byte)'<');
                    if (lt < 0 || lt >= rowEnd)
                    {
                        break; // no more cells before "</row" (or, on a truncated file, at all)
                    }
                    p = lt;
                    if (IsCellStart(buf, len, p))
                    {
                        p = ParseCellSpan(buf, len, p);
                    }
                    else if (!SkipMarkupSpan(buf, len, ref p))
                    {
                        break;
                    }
                }
                _pos = rowEnd;
                int gt = IndexOfBounded(buf, len, _pos, (byte)'>'); // already buffered — see EnsureRowBuffered
                _pos = gt < 0 ? len : gt + 1;
            }

            // Parses one <c>...</c> element starting at `p` (already known to be a cell — IsCellStart)
            // and returns the position right after it. Everything here operates on `buf`/`len`, which
            // ParseRow already guarantees cover the whole row (through "</row...>"'s closing '>'), so —
            // unlike the pre-T3.1 version — there is no Fill/compaction risk and therefore no need to
            // rewind `p` on a fast-path miss: the miss branch simply never advanced `p` past the open tag,
            // so falling through to the general "</c>" search below picks up from exactly where the fast
            // path started looking.
            private int ParseCellSpan(byte[] buf, int len, int p)
            {
                int gt = IndexOfBounded(buf, len, p, (byte)'>'); // end of the <c ...> open tag
                if (gt < 0)
                {
                    return len; // malformed: unclosed <c open tag within the buffered row
                }
                var header = ReadCellOpenTagSpan(buf, ref p, gt);
                if (header.SelfClose)
                {
                    return p; // empty cell — store nothing
                }

                // Fast path for the common bare-<v> shape (Number/Shared/Bool/Error, and a cached
                // t="str" formula result with no <f> element — see FormulaCellHasFormulaType). Raw '<'
                // can never appear inside valid XML text content, so the next '<' after "<v>" is
                // guaranteed to start "</v>" — one single-byte search replaces the general path's
                // "</c>" scan over the whole cell body followed by a second "<v>"/"</v>" scan inside it.
                // Inline-string and other formula shapes don't start with literal "<v>", so they
                // naturally fall through unchanged.
                if (buf.AsSpan(p, Math.Min(3, len - p)).StartsWith("<v>"u8))
                {
                    int valueStart = p + 3;
                    int lt = IndexOfBounded(buf, len, valueStart, (byte)'<');
                    if (lt >= 0 && buf.AsSpan(lt, Math.Min(8, len - lt)).StartsWith("</v></c>"u8))
                    {
                        ReadOnlySpan<byte> value = buf.AsSpan(valueStart, lt - valueStart);
                        if (header.Kind == Kind.Shared)
                        {
                            EmitShared(value, header.Col, header.Style);
                        }
                        else
                        {
                            EmitScalarValue(header.Kind, value, header.Col, header.Style);
                        }
                        return lt + 8;
                    }
                }

                ReadOnlySpan<byte> cClose = CClose; // "</c>", or "</x:c>" for a prefixed sheet
                int cEnd = IndexOfSeqBounded(buf, len, p, cClose); // ensures whole cell contiguous
                if (cEnd < 0)
                {
                    return len;
                }
                EmitCell(header.Kind, buf.AsSpan(p, cEnd - p), header.Col, header.Style);
                return cEnd + cClose.Length;
            }

            private enum HeadKind { End, Row, Skip }

            // Dispatches on the byte right after '<' before doing any StartsWith work, so the
            // overwhelmingly common "<row" case (and, per call site, the rare end-tags) each pay for
            // exactly one span comparison instead of up to three probes that mostly miss.
            private HeadKind ClassifyHead()
            {
                int avail = _len - _pos;
                if (avail < 2)
                {
                    return HeadKind.Skip;
                }
                if (_ns is not null)
                {
                    return ClassifyHeadPrefixed(_buf.AsSpan(_pos, avail));
                }
                switch (_buf[_pos + 1])
                {
                    case (byte)'r':
                        var rowHead = _buf.AsSpan(_pos, Math.Min(5, avail));
                        return rowHead.StartsWith("<row"u8) && (rowHead.Length < 5 || IsBoundary(rowHead[4]))
                            ? HeadKind.Row
                            : HeadKind.Skip;
                    case (byte)'/':
                        var endHead = _buf.AsSpan(_pos, Math.Min(11, avail));
                        return endHead.StartsWith("</sheetData"u8) || endHead.StartsWith("</worksheet"u8)
                            ? HeadKind.End
                            : HeadKind.Skip;
                    default:
                        return HeadKind.Skip;
                }
            }

            // Prefixed twin of ClassifyHead: matches "<x:row" (with a name boundary so it can't collide
            // with "<x:rowBreaks"), "</x:sheetData", "</x:worksheet". Runs once per top-level element,
            // never inside the per-cell loop, so field-token matching here costs nothing measurable.
            private HeadKind ClassifyHeadPrefixed(ReadOnlySpan<byte> head)
            {
                if (StartsWithElement(head, _ns!.RowOpen))
                {
                    return HeadKind.Row;
                }
                if (head.StartsWith(_ns.SheetDataEnd) || head.StartsWith(_ns.WorksheetEnd))
                {
                    return HeadKind.End;
                }
                return HeadKind.Skip;
            }

            // token is a full element open like "<x:row"; require a name-boundary byte after it so a
            // prefix match ("<x:row") doesn't swallow a longer sibling ("<x:rowBreaks").
            private static bool StartsWithElement(ReadOnlySpan<byte> span, ReadOnlySpan<byte> token)
            {
                return span.StartsWith(token) && (span.Length == token.Length || IsBoundary(span[token.Length]));
            }

            // "</row" itself is never reachable here — ParseRow's loop stops as soon as the found '<'
            // reaches `rowEnd` (exactly where "</row" starts) — so this only needs to tell a cell apart
            // from anything else (comments, CDATA, extension elements), which SkipMarkupSpan handles.
            private bool IsCellStart(byte[] buf, int len, int p)
            {
                int avail = len - p;
                if (_ns is not null)
                {
                    // The whole row is buffered here (EnsureRowBuffered), so the token can't be truncated.
                    return StartsWithElement(buf.AsSpan(p, avail), _ns.CellOpen);
                }
                if (avail < 2 || buf[p + 1] != (byte)'c')
                {
                    return false;
                }
                var cellHead = buf.AsSpan(p, Math.Min(3, avail));
                return cellHead.StartsWith("<c"u8) && (cellHead.Length < 3 || IsBoundary(cellHead[2]));
            }

            private readonly record struct CellHeader(int Col, int Style, Kind Kind, bool SelfClose);

            // Parses the <c ...> open tag ending at `gt`, advances `p` past it, and returns the extracted
            // (non-span) attributes.
            private CellHeader ReadCellOpenTagSpan(byte[] buf, ref int p, int gt)
            {
                var open = buf.AsSpan(p, gt - p + 1);
                ScanCellAttributes(open, out var rRef, out var sVal, out var tVal);

                // ColumnIndex returns -1 for a missing or malformed ref; fall back to the running column.
                int col = XlsxXml.ColumnIndex(rRef);
                if (col < 0)
                {
                    col = _nextCol;
                }
                _nextCol = col + 1;
                int style = ParseInt(sVal);
                var kind = ClassifyKind(tVal);
                bool selfClose = buf[gt - 1] == '/';
                p = gt + 1; // consume open tag; p now at inner start (or next element if self-closed)
                return new CellHeader(col, style, kind, selfClose);
            }

            private ValueTask<bool> BeginRowAsync()
            {
                int rel = _buf.AsSpan(_pos, _len - _pos).IndexOf((byte)'>');
                if (rel >= 0)
                {
                    return new ValueTask<bool>(BeginRowAt(_pos + rel));
                }
                return _eof ? new ValueTask<bool>(MissingRowOpenTag()) : BeginRowSlowAsync();
            }

            private async ValueTask<bool> BeginRowSlowAsync()
            {
                int gt = await IndexOfSlowAsync((byte)'>').ConfigureAwait(false);
                if (gt < 0)
                {
                    return MissingRowOpenTag();
                }
                return BeginRowAt(gt);
            }

            private bool BeginRowAt(int gt)
            {
                _pos = gt + 1;
                _acc.Reset();
                _nextCol = 0;
                return _buf[gt - 1] == '/';
            }

            private bool MissingRowOpenTag()
            {
                _pos = _len;
                _acc.Reset();
                _nextCol = 0;
                return true;
            }

            // Extracts the r/s/t attribute values from a `<c ...>` open tag in a single forward pass —
            // far cheaper than three separate IndexOf scans, whose per-call SIMD setup dominated for these
            // few-byte tags (and bare `<c>` number cells paid for three full misses). Any other attribute
            // is skipped. Returned spans alias `open`, so they live only as long as the caller's buffer.
            private static void ScanCellAttributes(
                ReadOnlySpan<byte> open,
                out ReadOnlySpan<byte> rRef,
                out ReadOnlySpan<byte> sVal,
                out ReadOnlySpan<byte> tVal)
            {
                rRef = sVal = tVal = default;
                int i = 2; // past "<c"
                while (i < open.Length && open[i] is not ((byte)'>' or (byte)'/'))
                {
                    if (!IsXmlSpace(open[i]))
                    {
                        i++;
                        continue;
                    }
                    i++; // consume the whitespace separating attributes

                    // Attribute name: runs up to '=' (bail out if this isn't a well-formed name="...").
                    int nameStart = i;
                    while (i < open.Length && open[i] is not ((byte)'=' or (byte)'>' or (byte)' '))
                    {
                        i++;
                    }
                    if (i >= open.Length || open[i] != (byte)'=')
                    {
                        continue;
                    }
                    int nameLen = i - nameStart;

                    // Attribute value: the run between the opening and closing quote.
                    i++; // '='
                    if (i >= open.Length || open[i] is not ((byte)'"' or (byte)'\''))
                    {
                        continue;
                    }
                    byte quote = open[i++];
                    int valueStart = i;
                    while (i < open.Length && open[i] != quote)
                    {
                        i++;
                    }
                    ReadOnlySpan<byte> value = open[valueStart..i];
                    i++; // closing quote

                    // Only the single-char attributes r/s/t matter; anything else is ignored.
                    if (nameLen != 1)
                    {
                        continue;
                    }
                    if (open[nameStart] == (byte)'r') { rRef = value; }
                    else if (open[nameStart] == (byte)'s') { sVal = value; }
                    else if (open[nameStart] == (byte)'t') { tVal = value; }
                }
            }

            private static bool IsXmlSpace(byte b)
            {
                return b is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n';
            }

            private enum Kind { Number, Shared, Inline, Bool, Error, Formula, IsoDate }

            // "" / "n" -> Number; "s" shared; "inlineStr" inline; "b" bool; "e" error; "str" formula
            // result; "d" ISO-8601 date (ECMA-376 §18.18.11 ST_CellType, written by some non-Excel producers).
            private static Kind ClassifyKind(ReadOnlySpan<byte> t)
            {
                return t.Length switch
                {
                    1 => t[0] switch
                    {
                        (byte)'s' => Kind.Shared,
                        (byte)'b' => Kind.Bool,
                        (byte)'e' => Kind.Error,
                        (byte)'d' => Kind.IsoDate,
                        _ => Kind.Number,
                    },
                    3 => Kind.Formula,   // "str"
                    9 => Kind.Inline,    // "inlineStr"
                    _ => Kind.Number,
                };
            }


            // Resolves a shared-string cell from its <v> index text. A non-numeric or negative index
            // (a corrupt or empty <v>) yields an empty string cell — never a silent substitution of
            // shared string 0, which is what parsing the garbage as index 0 used to produce.
            private void EmitShared(ReadOnlySpan<byte> indexText, int col, int style)
            {
                if (Utf8Parser.TryParse(indexText, out int index, out _) && index >= 0)
                {
                    var (start, len) = _reader.SharedAt(index);
                    _acc.Add(col, start, len, CellType.ExcelString, style, fromShared: true);
                    return;
                }
                _acc.Add(col, _acc.ValueLength, 0, CellType.ExcelString, style, fromShared: false);
            }

            private void EmitCell(Kind kind, ReadOnlySpan<byte> inner, int col, int style)
            {
                // Shared strings: <v> holds an index; point the cell at that slice of the shared buffer.
                if (kind == Kind.Shared)
                {
                    EmitShared(ElementText(inner, VOpen, VClose), col, style);
                    return;
                }

                if (kind == Kind.Inline)
                {
                    int vStart = _acc.ValueLength;
                    Span<byte> dst = _acc.ReserveValueSpan(inner.Length);
                    int written = _ns is null
                        ? XlsxXml.WriteTextRuns(inner, dst)
                        : XlsxXml.WriteTextRuns(inner, dst, _ns.TOpen, _ns.TClose, _ns.RPhOpen, _ns.RPhClose);
                    _acc.Advance(written);
                    _acc.Add(col, vStart, _acc.ValueLength - vStart, CellType.ExcelString, style, fromShared: false);
                    return;
                }

                EmitScalarValue(kind, ElementText(inner, VOpen, VClose), col, style);
            }

            // Handles every Kind whose content is bare "<v>...</v>" with no other wrapper: Number,
            // Bool, Error, and Formula (a cached t="str" result with no <f> element). Shared uses this
            // shape too but is handled by its callers directly, since its <v> holds a shared-string
            // index rather than the cell's own text. Shared by EmitCell (which locates `v` via
            // ElementText over the whole cell body) and ParseCell's fast path (which locates it via a
            // direct '<' search) — everything after the value text is found is identical either way.
            private void EmitScalarValue(Kind kind, ReadOnlySpan<byte> v, int col, int style)
            {
                // t="d": <v> holds ISO-8601 date text, not a serial. Parse it and store a 1900-system
                // serial so the cell behaves exactly like a style-based date cell (numeric, Type=Date).
                if (kind == Kind.IsoDate)
                {
                    EmitIsoDate(v, col, style);
                    return;
                }
                CellType cellType = kind switch
                {
                    Kind.Bool => CellType.Boolean,
                    Kind.Error => CellType.Error,
                    Kind.Formula => CellType.Formula,
                    _ => _reader.IsDateStyle(style) ? CellType.Date : CellType.Number,
                };
                int vStart = _acc.ValueLength;
                // Number/Bool/Error <v> text is pure ASCII digits/bool/error-code — it can never contain
                // an XML entity, so skip the decode scan. Only formula string results (t="str") can.
                if (kind == Kind.Formula)
                {
                    AppendDecoded(v);
                    _acc.Add(col, vStart, _acc.ValueLength - vStart, cellType, style, fromShared: false);
                    return;
                }
                AppendRaw(v);
                // Parse plain (non-exponent) numeric text at scan time so consumers (TryGetDouble/
                // TryParse<double>) skip the general double.TryParse round trip; the raw text is kept
                // either way so Value/GetString stay byte-identical. FastDouble.TryParse only accepts
                // inputs it can prove bit-identical to double.TryParse, so anything else (exponent form,
                // 17+ significant digits) just leaves hasNumber false and falls back at consume time.
                double number = 0;
                bool hasNumber = kind == Kind.Number && FastDouble.TryParse(v, out number);
                _acc.Add(col, vStart, _acc.ValueLength - vStart, cellType, style, fromShared: false,
                    number: number, hasNumber: hasNumber);
            }

            // Stores a t="d" ISO-8601 cell as a numeric Excel date serial (identical shape to a
            // style-based date cell), so TryGetDateTime works and GetString matches other date cells.
            private void EmitIsoDate(ReadOnlySpan<byte> v, int col, int style)
            {
                if (TryParseIsoDate(v, out DateTime dt))
                {
                    // ponytail: dt.ToOADate() is exact for dates >= 1900-03-01, which every real t="d"
                    // value is; pre-1900-03-01 would shift a day through the reader's 1900-leap fixup.
                    double serial = dt.ToOADate();
                    int start = _acc.ValueLength;
                    Span<byte> dst = _acc.ReserveValueSpan(32);
                    Utf8Formatter.TryFormat(serial, dst, out int written);
                    _acc.Advance(written);
                    _acc.Add(col, start, written, CellType.Date, style, fromShared: false, number: serial, hasNumber: true);
                    return;
                }
                // Unparseable ISO text: keep it verbatim as a string so nothing is silently dropped.
                int s = _acc.ValueLength;
                AppendRaw(v);
                _acc.Add(col, s, _acc.ValueLength - s, CellType.ExcelString, style, fromShared: false);
            }

            [SkipLocalsInit]
            private static bool TryParseIsoDate(ReadOnlySpan<byte> utf8, out DateTime value)
            {
                // ST_Xstring ISO dates are always ASCII; transcode to chars for DateTime.TryParse.
                if (utf8.Length is 0 or > 40)
                {
                    value = default;
                    return false;
                }
                Span<char> chars = stackalloc char[40];
                for (int i = 0; i < utf8.Length; i++)
                {
                    chars[i] = (char)utf8[i];
                }
                return DateTime.TryParse(chars[..utf8.Length], CultureInfo.InvariantCulture,
                    DateTimeStyles.RoundtripKind | DateTimeStyles.AllowWhiteSpaces, out value);
            }

            private static ReadOnlySpan<byte> ElementText(ReadOnlySpan<byte> inner, ReadOnlySpan<byte> openTag, ReadOnlySpan<byte> closeTag)
            {
                int s = inner.IndexOf(openTag);
                if (s < 0)
                {
                    return default;
                }
                s += openTag.Length;
                int e = inner[s..].IndexOf(closeTag);
                return e < 0 ? default : inner.Slice(s, e);
            }

            private void AppendDecoded(ReadOnlySpan<byte> src)
            {
                if (src.IsEmpty)
                {
                    return;
                }
                Span<byte> dst = _acc.ReserveValueSpan(src.Length);
                _acc.Advance(XlsxXml.Decode(src, dst));
            }

            // Copies verbatim, skipping the entity-decode scan — for <v> text known to never contain
            // XML entities (numbers, bools, error codes).
            private void AppendRaw(ReadOnlySpan<byte> src)
            {
                if (src.IsEmpty)
                {
                    return;
                }
                Span<byte> dst = _acc.ReserveValueSpan(src.Length);
                src.CopyTo(dst);
                _acc.Advance(src.Length);
            }

            private bool SkipMarkup()
            {
                Ensure(9);
                if (_buf.AsSpan(_pos, Math.Min(4, _len - _pos)).StartsWith("<!--"u8))
                {
                    int end = IndexOfSeq("-->"u8);
                    _pos = end < 0 ? _len : end + 3;
                    return end >= 0;
                }
                if (_buf.AsSpan(_pos, Math.Min(9, _len - _pos)).StartsWith("<![CDATA["u8))
                {
                    int end = IndexOfSeq("]]>"u8);
                    _pos = end < 0 ? _len : end + 3;
                    return end >= 0;
                }

                int skip = IndexOf((byte)'>');
                if (skip < 0)
                {
                    return false;
                }
                _pos = skip + 1;
                return true;
            }

            private async ValueTask<bool> SkipMarkupAsync()
            {
                await EnsureAsync(9).ConfigureAwait(false);
                if (_buf.AsSpan(_pos, Math.Min(4, _len - _pos)).StartsWith("<!--"u8))
                {
                    int end = await IndexOfSeqAsync(MarkupSeq.CommentEnd).ConfigureAwait(false);
                    _pos = end < 0 ? _len : end + 3;
                    return end >= 0;
                }
                if (_buf.AsSpan(_pos, Math.Min(9, _len - _pos)).StartsWith("<![CDATA["u8))
                {
                    int end = await IndexOfSeqAsync(MarkupSeq.CDataEnd).ConfigureAwait(false);
                    _pos = end < 0 ? _len : end + 3;
                    return end >= 0;
                }

                int skip = await IndexOfAsync((byte)'>').ConfigureAwait(false);
                if (skip < 0)
                {
                    return false;
                }
                _pos = skip + 1;
                return true;
            }

            // ParseRow's counterpart to SkipMarkup: same shape, but bounded to the already-fully-buffered
            // `buf[0..len)` window via a local cursor instead of `_pos`/Ensure/Fill.
            private static bool SkipMarkupSpan(byte[] buf, int len, ref int p)
            {
                if (buf.AsSpan(p, Math.Min(4, len - p)).StartsWith("<!--"u8))
                {
                    int end = IndexOfSeqBounded(buf, len, p, "-->"u8);
                    p = end < 0 ? len : end + 3;
                    return end >= 0;
                }
                if (buf.AsSpan(p, Math.Min(9, len - p)).StartsWith("<![CDATA["u8))
                {
                    int end = IndexOfSeqBounded(buf, len, p, "]]>"u8);
                    p = end < 0 ? len : end + 3;
                    return end >= 0;
                }

                int skip = IndexOfBounded(buf, len, p, (byte)'>');
                if (skip < 0)
                {
                    return false;
                }
                p = skip + 1;
                return true;
            }

            private static int ParseInt(ReadOnlySpan<byte> src)
            {
                return Utf8Parser.TryParse(src, out int v, out _) ? v : 0;
            }

            private static bool IsBoundary(byte b)
            {
                return b is (byte)' ' or (byte)'>' or (byte)'/' or (byte)'\t' or (byte)'\r' or (byte)'\n';
            }

            // --- buffer management ---
            // After every Fill the window [_pos.._len) is rescanned from the start; that re-reads a few
            // bytes but keeps the search loops trivial and handles delimiters split across a refill for free.

            private int IndexOf(byte b)
            {
                while (true)
                {
                    int rel = _buf.AsSpan(_pos, _len - _pos).IndexOf(b);
                    if (rel >= 0)
                    {
                        return _pos + rel;
                    }
                    if (_eof)
                    {
                        return -1;
                    }
                    _io.Fill(_sheet!);
                }
            }

            private int IndexOfSeq(ReadOnlySpan<byte> seq)
            {
                while (true)
                {
                    int rel = _buf.AsSpan(_pos, _len - _pos).IndexOf(seq);
                    if (rel >= 0)
                    {
                        return _pos + rel;
                    }
                    if (_eof)
                    {
                        return -1;
                    }
                    _io.Fill(_sheet!);
                }
            }

            // Bounded, Fill-free counterparts used by ParseRow/ParseCellSpan/SkipMarkupSpan once
            // EnsureRowBuffered(Async) has already guaranteed the whole row is buffered — a plain span
            // search, never able to trigger PrepareBuffer compaction, so `from`/the returned index stay
            // valid for as long as the caller doesn't call anything that can Fill in between.
            private static int IndexOfBounded(byte[] buf, int boundExclusive, int from, byte b)
            {
                int rel = buf.AsSpan(from, boundExclusive - from).IndexOf(b);
                return rel < 0 ? -1 : from + rel;
            }

            private static int IndexOfSeqBounded(byte[] buf, int boundExclusive, int from, ReadOnlySpan<byte> seq)
            {
                int rel = buf.AsSpan(from, boundExclusive - from).IndexOf(seq);
                return rel < 0 ? -1 : from + rel;
            }

            private void Ensure(int n)
            {
                _io.Ensure(_sheet!, n);
            }

            // Async twins of the search/refill primitives. Each is split so the common case (the target is
            // already in the buffered window) returns a completed task with no async state machine, and
            // only a real refill on a buffer miss takes the awaiting slow path. No span crosses an await.

            private ValueTask<int> IndexOfAsync(byte b)
            {
                int rel = _buf.AsSpan(_pos, _len - _pos).IndexOf(b);
                if (rel >= 0)
                {
                    return new ValueTask<int>(_pos + rel);
                }
                return _eof ? new ValueTask<int>(-1) : IndexOfSlowAsync(b);
            }

            private async ValueTask<int> IndexOfSlowAsync(byte b)
            {
                do
                {
                    await FillAsync().ConfigureAwait(false);
                    int rel = _buf.AsSpan(_pos, _len - _pos).IndexOf(b);
                    if (rel >= 0)
                    {
                        return _pos + rel;
                    }
                }
                while (!_eof);
                return -1;
            }

            // Grows the buffer (blocking Fill loop) until the whole row — cell data plus "</row...>"'s
            // closing '>' — is present, so ParseRow can run entirely Fill-free over a local span+cursor.
            // Returns the absolute position of "</row"'s '<' (the row's cell-data end); on a truncated
            // file (no </row> before EOF) returns _len instead, so ParseRow still parses whatever cells
            // are present before bailing, matching the original per-cell path's truncation behavior.
            private int EnsureRowBuffered()
            {
                while (true)
                {
                    int rowEnd = FindSeq(MarkupSeq.RowEnd, _pos);
                    if (rowEnd >= 0 && _buf.AsSpan(rowEnd, _len - rowEnd).IndexOf((byte)'>') >= 0)
                    {
                        return rowEnd;
                    }
                    if (_eof)
                    {
                        return _len;
                    }
                    _io.Fill(_sheet!);
                }
            }

            // Async twin: non-async fast path checks synchronously first (the buffer already has the
            // whole row ~99.9% of the time), only awaiting a Fill on a genuine miss.
            private ValueTask<int> EnsureRowBufferedAsync()
            {
                int rowEnd = FindSeq(MarkupSeq.RowEnd, _pos);
                if (rowEnd >= 0 && _buf.AsSpan(rowEnd, _len - rowEnd).IndexOf((byte)'>') >= 0)
                {
                    return new ValueTask<int>(rowEnd);
                }
                return _eof ? new ValueTask<int>(_len) : EnsureRowBufferedSlowAsync();
            }

            private async ValueTask<int> EnsureRowBufferedSlowAsync()
            {
                do
                {
                    await FillAsync().ConfigureAwait(false);
                    int rowEnd = FindSeq(MarkupSeq.RowEnd, _pos);
                    if (rowEnd >= 0 && _buf.AsSpan(rowEnd, _len - rowEnd).IndexOf((byte)'>') >= 0)
                    {
                        return rowEnd;
                    }
                }
                while (!_eof);
                return _len;
            }

            private enum MarkupSeq { CommentEnd, CDataEnd, RowEnd }

            private int FindSeq(MarkupSeq seq, int start)
            {
                int rel = seq switch
                {
                    MarkupSeq.CommentEnd => _buf.AsSpan(start, _len - start).IndexOf("-->"u8),
                    MarkupSeq.CDataEnd => _buf.AsSpan(start, _len - start).IndexOf("]]>"u8),
                    _ => _buf.AsSpan(start, _len - start).IndexOf(_ns is null ? "</row"u8 : _ns.RowEnd),
                };
                return rel < 0 ? -1 : start + rel;
            }

            private ValueTask<int> IndexOfSeqAsync(MarkupSeq seq)
            {
                int index = FindSeq(seq, _pos);
                if (index >= 0)
                {
                    return new ValueTask<int>(index);
                }
                return _eof ? new ValueTask<int>(-1) : IndexOfSeqFromAsync(seq);
            }

            // PrepareBuffer compacts the window on every fill, so an absolute index captured before a
            // fill is invalid afterward. Rescan the retained window from _pos each time — this also
            // catches a sequence split across the previous buffer boundary.
            private async ValueTask<int> IndexOfSeqFromAsync(MarkupSeq seq)
            {
                do
                {
                    await FillAsync().ConfigureAwait(false);
                    int index = FindSeq(seq, _pos);
                    if (index >= 0)
                    {
                        return index;
                    }
                }
                while (!_eof);
                return -1;
            }

            private ValueTask EnsureAsync(int n)
            {
                return _io.EnsureAsync(_sheet!, n, _ct);
            }

            private ValueTask FillAsync()
            {
                return _io.FillAsync(_sheet!, _ct);
            }

            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
                Justification = "_sheet is opened for this enumerator and owned by it.")]
            public void Dispose()
            {
                _sheet?.Dispose();
                _sheet = null;
                ReturnBuffers();
            }

            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
                Justification = "_sheet is opened for this enumerator and owned by it.")]
            public async ValueTask DisposeAsync()
            {
                if (_sheet is not null)
                {
                    await _sheet.DisposeAsync().ConfigureAwait(false);
                    _sheet = null;
                }
                ReturnBuffers();
            }

            private void ReturnBuffers()
            {
                _io.Return();
                _acc.Return();
            }
        }
    }
}
