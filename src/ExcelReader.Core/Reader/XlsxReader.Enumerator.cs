using System.Buffers.Text;
using System.Diagnostics.CodeAnalysis;
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

            internal Enumerator(XlsxReader reader, Stream sheet, CancellationToken ct = default)
            {
                _reader = reader;
                _sheet = sheet;
                _ct = ct;
                _io = new BufferedStreamCursor(reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes));
                _acc = new CellAccumulator(reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes));
            }

            public Row Current =>
                new(_acc.CellSpan, _acc.ValueSpan, _reader.SharedSpan);

            // The sync and async row scanners share every span-touching helper below (ClassifyHead,
            // ClassifyRowHead, ReadCellOpenTag, EmitCell). The two families differ only in whether the
            // buffer refill (Fill / FillAsync) and the byte searches await — so the span work is factored
            // into sync helpers that never hold a span across an await.

            public bool MoveNext()
            {
                while (true)
                {
                    // Fast path: in compact output _pos already sits on the next '<', so skip the scan.
                    int lt = _pos < _len && _buf[_pos] == (byte)'<' ? _pos : IndexOf((byte)'<');
                    if (lt < 0)
                    {
                        return false;
                    }
                    _pos = lt;
                    Ensure(12);
                    switch (ClassifyHead())
                    {
                        case HeadKind.End:
                            return false;
                        case HeadKind.Row:
                            if (!BeginRow())
                            {
                                ParseRow();
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

                    ValueTask ensureTask = EnsureAsync(12);
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
                                    ValueTask rowBufTask = EnsureRowBufferedAsync();
                                    if (!rowBufTask.IsCompletedSuccessfully)
                                    {
                                        return FinishRowAfterAsync(rowBufTask);
                                    }
                                    ParseRow();
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
            private async ValueTask<bool> FinishRowAfterAsync(ValueTask pendingRowBuffered)
            {
                await pendingRowBuffered.ConfigureAwait(false);
                ParseRow();
                return true;
            }

            // Consumes the <row ...> open tag and resets per-row state. Call only after ClassifyHead()==Row.
            private bool BeginRow()
            {
                int gt = IndexOf((byte)'>'); // open tag already fully buffered by the Ensure(12) above
                return BeginRowAt(gt);
            }

            private void ParseRow()
            {
                while (true)
                {
                    int lt = _pos < _len && _buf[_pos] == (byte)'<' ? _pos : IndexOf((byte)'<');
                    if (lt < 0)
                    {
                        return;
                    }
                    _pos = lt;
                    Ensure(6);
                    switch (ClassifyRowHead())
                    {
                        case RowHead.EndRow:
                            int gt = IndexOf((byte)'>');
                            _pos = gt < 0 ? _len : gt + 1;
                            return;
                        case RowHead.Cell:
                            ParseCell();
                            break;
                        default:
                            if (!SkipMarkup())
                            {
                                return;
                            }
                            break;
                    }
                }
            }

            private void ParseCell()
            {
                int gt = IndexOf((byte)'>'); // end of the <c ...> open tag (buffered, no shift yet)
                var header = ReadCellOpenTag(gt);
                if (header.SelfClose)
                {
                    return; // empty cell — store nothing; _pos already past the tag
                }

                // Fast path for the common bare-<v> shape (Number/Shared/Bool/Error, and a cached
                // t="str" formula result with no <f> element — see FormulaCellHasFormulaType). Raw '<'
                // can never appear inside valid XML text content, so the next '<' after "<v>" is
                // guaranteed to start "</v>" — one single-byte search replaces the general path's
                // "</c>" scan over the whole cell body followed by a second "<v>"/"</v>" scan inside it.
                // Inline-string and other formula shapes don't start with literal "<v>", so they
                // naturally fall through unchanged. `valueLen` (a relative distance, not a captured
                // absolute index) keeps this correct across any Fill/compaction inside IndexOf/Ensure —
                // see the buffer-management notes below.
                Ensure(3);
                if (_buf.AsSpan(_pos, Math.Min(3, _len - _pos)).StartsWith("<v>"u8))
                {
                    _pos += 3;
                    int lt = IndexOf((byte)'<');
                    if (lt >= 0)
                    {
                        int valueLen = lt - _pos;
                        _pos = lt;
                        Ensure(8); // "</v></c>"
                        if (_buf.AsSpan(_pos, Math.Min(8, _len - _pos)).StartsWith("</v></c>"u8))
                        {
                            ReadOnlySpan<byte> value = _buf.AsSpan(_pos - valueLen, valueLen);
                            if (header.Kind == Kind.Shared)
                            {
                                var (start, len) = _reader.SharedAt(ParseInt(value));
                                _acc.Add(header.Col, start, len, CellType.ExcelString, header.Style, fromShared: true);
                            }
                            else
                            {
                                EmitScalarValue(header.Kind, value, header.Col, header.Style);
                            }
                            _pos += 8;
                            return;
                        }
                        _pos -= valueLen; // not the expected shape — rewind to the value's start
                    }
                    else
                    {
                        _pos -= 3; // no closing '<' found at all — rewind past "<v>"
                    }
                }

                int cEnd = IndexOfSeq("</c>"u8); // ensures whole cell contiguous; shifts _pos consistently
                if (cEnd < 0)
                {
                    _pos = _len;
                    return;
                }
                EmitCell(header.Kind, _buf.AsSpan(_pos, cEnd - _pos), header.Col, header.Style);
                _pos = cEnd + 4;
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

            private enum RowHead { EndRow, Cell, Other }

            private RowHead ClassifyRowHead()
            {
                int avail = _len - _pos;
                if (avail < 2)
                {
                    return RowHead.Other;
                }
                switch (_buf[_pos + 1])
                {
                    case (byte)'c':
                        var cellHead = _buf.AsSpan(_pos, Math.Min(3, avail));
                        return cellHead.StartsWith("<c"u8) && (cellHead.Length < 3 || IsBoundary(cellHead[2]))
                            ? RowHead.Cell
                            : RowHead.Other;
                    case (byte)'/':
                        var endHead = _buf.AsSpan(_pos, Math.Min(5, avail));
                        return endHead.StartsWith("</row"u8) ? RowHead.EndRow : RowHead.Other;
                    default:
                        return RowHead.Other;
                }
            }

            private readonly record struct CellHeader(int Col, int Style, Kind Kind, bool SelfClose);

            // Parses the <c ...> open tag ending at `gt`, advances _pos past it, and returns the extracted
            // (non-span) attributes. Holds spans only within this synchronous call, so callers may await after.
            private CellHeader ReadCellOpenTag(int gt)
            {
                var open = _buf.AsSpan(_pos, gt - _pos + 1);
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
                bool selfClose = _buf[gt - 1] == '/';
                _pos = gt + 1; // consume open tag; _pos now at inner start (or next element if self-closed)
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

            private enum Kind { Number, Shared, Inline, Bool, Error, Formula }

            // "" / "n" -> Number; "s" shared; "inlineStr" inline; "b" bool; "e" error; "str" formula result.
            private static Kind ClassifyKind(ReadOnlySpan<byte> t)
            {
                return t.Length switch
                {
                    1 => t[0] switch
                    {
                        (byte)'s' => Kind.Shared,
                        (byte)'b' => Kind.Bool,
                        (byte)'e' => Kind.Error,
                        _ => Kind.Number,
                    },
                    3 => Kind.Formula,   // "str"
                    9 => Kind.Inline,    // "inlineStr"
                    _ => Kind.Number,
                };
            }


            private void EmitCell(Kind kind, ReadOnlySpan<byte> inner, int col, int style)
            {
                // Shared strings: <v> holds an index; point the cell at that slice of the shared buffer.
                if (kind == Kind.Shared)
                {
                    var (start, len) = _reader.SharedAt(ParseInt(ElementText(inner, "<v>"u8, "</v>"u8)));
                    _acc.Add(col, start, len, CellType.ExcelString, style, fromShared: true);
                    return;
                }

                if (kind == Kind.Inline)
                {
                    int vStart = _acc.ValueLength;
                    Span<byte> dst = _acc.ReserveValueSpan(inner.Length);
                    _acc.Advance(XlsxXml.WriteTextRuns(inner, dst));
                    _acc.Add(col, vStart, _acc.ValueLength - vStart, CellType.ExcelString, style, fromShared: false);
                    return;
                }

                EmitScalarValue(kind, ElementText(inner, "<v>"u8, "</v>"u8), col, style);
            }

            // Handles every Kind whose content is bare "<v>...</v>" with no other wrapper: Number,
            // Bool, Error, and Formula (a cached t="str" result with no <f> element). Shared uses this
            // shape too but is handled by its callers directly, since its <v> holds a shared-string
            // index rather than the cell's own text. Shared by EmitCell (which locates `v` via
            // ElementText over the whole cell body) and ParseCell's fast path (which locates it via a
            // direct '<' search) — everything after the value text is found is identical either way.
            private void EmitScalarValue(Kind kind, ReadOnlySpan<byte> v, int col, int style)
            {
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

            private ValueTask EnsureRowBufferedAsync()
            {
                return IsRowBuffered() ? ValueTask.CompletedTask : EnsureRowBufferedSlowAsync();
            }

            private bool IsRowBuffered()
            {
                int rowEnd = FindSeq(MarkupSeq.RowEnd, _pos);
                if (rowEnd < 0)
                {
                    if (_eof)
                    {
                        _pos = _len;
                        return true;
                    }
                    return false;
                }

                if (_buf.AsSpan(rowEnd, _len - rowEnd).IndexOf((byte)'>') >= 0)
                {
                    return true;
                }
                if (_eof)
                {
                    _pos = _len;
                    return true;
                }
                return false;
            }

            private async ValueTask EnsureRowBufferedSlowAsync()
            {
                int rowEnd = await IndexOfSeqFromAsync(MarkupSeq.RowEnd).ConfigureAwait(false);
                if (rowEnd < 0)
                {
                    _pos = _len;
                    return;
                }
                int gt = await IndexOfFromAsync((byte)'>', rowEnd).ConfigureAwait(false);
                if (gt < 0)
                {
                    _pos = _len;
                }
            }

            private enum MarkupSeq { CommentEnd, CDataEnd, RowEnd }

            private int FindSeq(MarkupSeq seq, int start)
            {
                int rel = seq switch
                {
                    MarkupSeq.CommentEnd => _buf.AsSpan(start, _len - start).IndexOf("-->"u8),
                    MarkupSeq.CDataEnd => _buf.AsSpan(start, _len - start).IndexOf("]]>"u8),
                    _ => _buf.AsSpan(start, _len - start).IndexOf("</row"u8),
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

            private async ValueTask<int> IndexOfFromAsync(byte b, int start)
            {
                int search = Math.Max(start, _pos);
                while (true)
                {
                    int rel = _buf.AsSpan(search, _len - search).IndexOf(b);
                    if (rel >= 0)
                    {
                        return search + rel;
                    }
                    if (_eof)
                    {
                        return -1;
                    }
                    int shift = _pos;
                    await FillAsync().ConfigureAwait(false);
                    search = shift > 0 ? Math.Max(_pos, search - shift) : search;
                }
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
