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
        /// <summary>Forward-only enumerator over an <see cref="XlsxReader"/> sheet's rows.</summary>
        /// <remarks>Low-memory: streams the sheet through a refillable pooled buffer, growing it as needed so a single <c>&lt;c&gt;...&lt;/c&gt;</c> element is always guaranteed contiguous.</remarks>
        [SuppressMessage("Design", "CA1034:Nested types should not be visible",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public sealed class Enumerator : PooledStreamRowEnumerator, IExcelRowEnumerator
        {
            [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed", Justification = "Borrowed, not owned.")]
            private readonly XlsxReader _reader;
            // Hoisted out of _reader to avoid a dependent load on the per-cell hot path.
            private readonly bool[] _styleIsDate;
            private readonly int[] _sharedOffsets;
            // Content-keyed dedup cache for inline/formula-string cells (see ExcelReaderOptions.InternStrings).
            private readonly Utf8StringCache? _contentCache;
            private int _nextCol;

            // Non-null only when this sheet's elements carry a namespace prefix (e.g. <x:row>), holding
            // the prefixed forms of every token the scanner matches. Detected once, lazily.
            private NsTokens? _ns;
            private bool _nsChecked;

            private ReadOnlySpan<byte> VOpen => _ns is null ? "<v>"u8 : _ns.VOpen;
            private ReadOnlySpan<byte> VClose => _ns is null ? "</v>"u8 : _ns.VClose;
            private ReadOnlySpan<byte> CClose => _ns is null ? "</c>"u8 : _ns.CClose;

            internal Enumerator(XlsxReader reader, Stream sheet, long entryLength = 0, CancellationToken ct = default)
                : base(sheet, reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes), WorkbookLookups.InitialBufferCapacity(entryLength), ownsSource: true, ct)
            {
                _reader = reader;
                _styleIsDate = reader._styleIsDate;
                _sharedOffsets = reader._sharedOffsets;
                _contentCache = reader._options.InternStrings ? new Utf8StringCache() : null;
            }

            /// <inheritdoc/>
            public Row Current =>
                new(_acc.CellSpan, _acc.ValueSpan, _reader.SharedSpan, _buf.AsSpan(0, _len), _reader.SharedStringCache, _contentCache);

            // Top-level scanning differs between sync and async only in whether the buffer refill
            // awaits, so that span work stays in sync helpers that never hold a span across an await.
            // Once inside a row, EnsureRowBuffered(Async) guarantees the whole row is buffered first,
            // so ParseRow/ParseCellSpan/EmitCell are shared by both paths unchanged.

            /// <inheritdoc/>
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

            // Returns a completed ValueTask when every step resolves synchronously, only falling to an
            // awaiting continuation at the exact step that needs a refill.
            /// <inheritdoc/>
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
                            return ReadRowAsync();
                        default:
                            ValueTask<bool>? skipResult = SkipMarkupOrContinue();
                            if (skipResult is null)
                            {
                                break; // markup skipped — continue scanning for the next element
                            }
                            return skipResult.Value;
                    }
                }
            }

            [SuppressMessage("VisualStudio.Threading", "VSTHRD103:Result synchronously blocks",
                Justification = "Every .Result access is guarded by IsCompletedSuccessfully immediately above it — never blocks.")]
            private ValueTask<bool> ReadRowAsync()
            {
                ValueTask<bool> beginTask = BeginRowAsync();
                if (!beginTask.IsCompletedSuccessfully)
                {
                    return AwaitThenRestartAsync(beginTask);
                }
                if (beginTask.Result)
                {
                    return new ValueTask<bool>(true);
                }

                ValueTask<int> rowBufferTask = EnsureRowBufferedAsync();
                if (!rowBufferTask.IsCompletedSuccessfully)
                {
                    return FinishRowAfterAsync(rowBufferTask);
                }
                ParseRow(rowBufferTask.Result);
                return new ValueTask<bool>(true);
            }

            // Returns null only when markup was skipped and enumeration should continue immediately.
            [SuppressMessage("VisualStudio.Threading", "VSTHRD002:Avoid problematic synchronous waits",
                Justification = "The .Result access is guarded by IsCompletedSuccessfully immediately above it — never blocks.")]
            [SuppressMessage("Reliability", "CA2012:Use ValueTasks correctly",
                Justification = "The ValueTask is either returned through AwaitThenRestartAsync or consumed once after confirming synchronous completion.")]
            private ValueTask<bool>? SkipMarkupOrContinue()
            {
                ValueTask<bool> skipTask = SkipMarkupAsync();
                if (!skipTask.IsCompletedSuccessfully)
                {
                    return AwaitThenRestartAsync(skipTask);
                }
                if (skipTask.Result)
                {
                    return null; // markup skipped — caller continues the scan loop
                }
                return new ValueTask<bool>(false); // end of sheetData/worksheet
            }

            // Safe to re-enter from the top: none of the pending steps this restarts commit a position
            // change until they resolve, so once the fill completes the (now-buffered) work just redoes.
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

            // Unlike AwaitThenRestartAsync, must not re-enter MoveNextAsync from the top — the row is
            // already open and that would misread its first cell as a new top-level element.
            private async ValueTask<bool> FinishRowAfterAsync(ValueTask<int> pendingRowBuffered)
            {
                int rowEnd = await pendingRowBuffered.ConfigureAwait(false);
                ParseRow(rowEnd);
                return true;
            }

            // Detects the sheet's element-name prefix (e.g. "x:" in <x:worksheet>) once, from the root
            // element at the start of the stream. Prefixed worksheets are rare.
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
                int gt = IndexOf((byte)'>');
                if (gt < 0)
                {
                    return MissingRowOpenTag();
                }
                return BeginRowAt(gt);
            }

            // `rowEnd` (the '<' starting "</row") is supplied by EnsureRowBuffered(Async), which already
            // grew the buffer until the whole row is present — so everything below is pure
            // ReadOnlySpan<byte> work with no Ensure/Fill and no mid-row compaction risk. `_pos` is
            // written back exactly once, after the whole row is consumed.
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
                int gt = IndexOfBounded(buf, len, _pos, (byte)'>');
                _pos = gt < 0 ? len : gt + 1;
            }

            // Parses one <c>...</c> element starting at `p` (already known to be a cell) and returns the
            // position right after it. `buf`/`len` cover the whole row, so a fast-path miss can safely
            // fall through to the general "</c>" search without rewinding `p`.
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

                // Fast path for the common bare-<v> shape: raw '<' can never appear inside valid XML
                // text content, so the next '<' after "<v>" is guaranteed to start "</v>" — one
                // single-byte search instead of a "</c>" scan plus a nested "<v>"/"</v>" scan.
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
                            EmitScalarValueFast(header.Kind, value, valueStart, header.Col, header.Style);
                        }
                        return lt + 8;
                    }
                }

                ReadOnlySpan<byte> cClose = CClose;
                int cEnd = IndexOfSeqBounded(buf, len, p, cClose);
                if (cEnd < 0)
                {
                    return len;
                }
                EmitCell(header.Kind, buf.AsSpan(p, cEnd - p), header.Col, header.Style);
                return cEnd + cClose.Length;
            }

            private enum HeadKind { End, Row, Skip }

            // Dispatches on the byte right after '<' before any StartsWith work, so the common "<row"
            // case costs one span comparison instead of several mostly-missing probes.
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

            // Prefixed twin of ClassifyHead. Runs once per top-level element, never per cell.
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

            // Requires a name-boundary byte after `token` so "<x:row" doesn't swallow "<x:rowBreaks".
            private static bool StartsWithElement(ReadOnlySpan<byte> span, ReadOnlySpan<byte> token)
            {
                return span.StartsWith(token) && (span.Length == token.Length || IsBoundary(span[token.Length]));
            }

            private bool IsCellStart(byte[] buf, int len, int p)
            {
                int avail = len - p;
                if (_ns is not null)
                {
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
                int style = XlsxXml.ParseIntOr(sVal, 0);
                var kind = ClassifyKind(tVal);
                bool selfClose = buf[gt - 1] == '/';
                p = gt + 1;
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

            // Extracts the r/s/t attribute values from a `<c ...>` open tag in one forward pass rather
            // than three separate IndexOf scans. Returned spans alias `open`.
            private static void ScanCellAttributes(
                ReadOnlySpan<byte> open,
                out ReadOnlySpan<byte> rRef,
                out ReadOnlySpan<byte> sVal,
                out ReadOnlySpan<byte> tVal)
            {
                rRef = sVal = tVal = default;
                int i = 2;
                while (i < open.Length && open[i] is not ((byte)'>' or (byte)'/'))
                {
                    if (!IsXmlSpace(open[i]))
                    {
                        i++;
                        continue;
                    }
                    i++;

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

                    i++;
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
                    i++;

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
            // result; "d" ISO-8601 date (written by some non-Excel producers).
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

            // A non-numeric or negative index yields an empty string cell, never a silent substitution
            // of shared string 0.
            private void EmitShared(ReadOnlySpan<byte> indexText, int col, int style)
            {
                if (Utf8Parser.TryParse(indexText, out int index, out _) && index >= 0)
                {
                    var (start, len, sharedIndex) = WorkbookLookups.SharedAt(_sharedOffsets, index);
                    _acc.Add(col, start, len, CellType.ExcelString, style, CellValueSource.Shared, sharedIndex: sharedIndex);
                    return;
                }
                _acc.Add(col, _acc.ValueLength, 0, CellType.ExcelString, style, CellValueSource.RowValues);
            }

            private void EmitCell(Kind kind, ReadOnlySpan<byte> inner, int col, int style)
            {
                if (kind == Kind.Shared)
                {
                    EmitShared(ElementText(inner, VOpen, VClose), col, style);
                    return;
                }

                if (kind != Kind.Inline)
                {
                    EmitScalarValue(kind, ElementText(inner, VOpen, VClose), col, style);
                    return;
                }

                int vStart = _acc.ValueLength;
                Span<byte> dst = _acc.ReserveValueSpan(inner.Length);
                int written = _ns is null
                    ? XlsxXml.WriteTextRuns(inner, dst)
                    : XlsxXml.WriteTextRuns(inner, dst, _ns.TOpen, _ns.TClose, _ns.RPhOpen, _ns.RPhClose);
                _acc.Advance(written);
                _acc.Add(col, vStart, _acc.ValueLength - vStart, CellType.ExcelString, style, CellValueSource.RowValues);
            }

            // Handles every Kind whose content is bare "<v>...</v>": Number, Bool, Error, Formula.
            private void EmitScalarValue(Kind kind, ReadOnlySpan<byte> v, int col, int style)
            {
                // t="d": <v> holds ISO-8601 date text, not a serial; store a 1900-system serial so the
                // cell behaves like a style-based date cell.
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
                    _ => WorkbookLookups.IsDateStyle(_styleIsDate, style) ? CellType.Date : CellType.Number,
                };
                int vStart = _acc.ValueLength;
                // Number/Bool/Error <v> text can never contain an XML entity; only formula results can.
                if (kind == Kind.Formula)
                {
                    AppendDecoded(v);
                    _acc.Add(col, vStart, _acc.ValueLength - vStart, cellType, style, CellValueSource.RowValues);
                    return;
                }
                AppendRaw(v);
                // FastDouble.TryParse only accepts inputs bit-identical to double.TryParse; anything
                // else leaves hasNumber false and falls back at consume time.
                double number = 0;
                bool hasNumber = kind == Kind.Number && FastDouble.TryParse(v, out number);
                _acc.Add(col, vStart, _acc.ValueLength - vStart, cellType, style, CellValueSource.RowValues,
                    number: number, hasNumber: hasNumber);
            }

            // Fast-path counterpart to EmitScalarValue: aliases `buf` directly instead of copying into
            // the accumulator. Falls back to EmitScalarValue for IsoDate (always reformats) or a
            // Formula result containing '&' (needs entity decoding).
            private void EmitScalarValueFast(Kind kind, ReadOnlySpan<byte> v, int valueStart, int col, int style)
            {
                if (kind == Kind.IsoDate || (kind == Kind.Formula && v.IndexOf((byte)'&') >= 0))
                {
                    EmitScalarValue(kind, v, col, style);
                    return;
                }
                CellType cellType = kind switch
                {
                    Kind.Bool => CellType.Boolean,
                    Kind.Error => CellType.Error,
                    Kind.Formula => CellType.Formula,
                    _ => WorkbookLookups.IsDateStyle(_styleIsDate, style) ? CellType.Date : CellType.Number,
                };
                double number = 0;
                bool hasNumber = kind == Kind.Number && FastDouble.TryParse(v, out number);
                _acc.Add(col, valueStart, v.Length, cellType, style, CellValueSource.RowBuffer,
                    number: number, hasNumber: hasNumber);
            }

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
                    _acc.Add(col, start, written, CellType.Date, style, CellValueSource.RowValues, number: serial, hasNumber: true);
                    return;
                }
                int s = _acc.ValueLength;
                AppendRaw(v);
                _acc.Add(col, s, _acc.ValueLength - s, CellType.ExcelString, style, CellValueSource.RowValues);
            }

            [SkipLocalsInit]
            private static bool TryParseIsoDate(ReadOnlySpan<byte> utf8, out DateTime value)
            {
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
                const DateTimeStyles dateParseFlag = DateTimeStyles.RoundtripKind | DateTimeStyles.AllowWhiteSpaces;
                if (!DateTime.TryParse(chars[..utf8.Length], CultureInfo.InvariantCulture, dateParseFlag, out value))
                {
                    return false;
                }
                // ToOADate throws below year 100; treat that as unparseable rather than crash the read.
                if (value.Year < 100)
                {
                    value = default;
                    return false;
                }
                return true;
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

            // Copies verbatim, skipping the entity-decode scan.
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

            // ParseRow's counterpart to SkipMarkup, bounded to the buffered window via a local cursor.
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

            private static bool IsBoundary(byte b)
            {
                return b is (byte)' ' or (byte)'>' or (byte)'/' or (byte)'\t' or (byte)'\r' or (byte)'\n';
            }

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
                    Fill();
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
                    Fill();
                }
            }

            // Bounded, Fill-free counterparts used once EnsureRowBuffered(Async) has already
            // guaranteed the whole row is buffered.
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

            // Grows the buffer until the whole row (through "</row...>"'s closing '>') is present.
            // Returns _len instead on a truncated file, so ParseRow still parses whatever is present.
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
                    Fill();
                }
            }

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

            // Rescans the retained window from _pos each time, since compaction invalidates an
            // absolute index captured before the fill.
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

        }
    }
}
