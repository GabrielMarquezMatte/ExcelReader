using System.Buffers.Text;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Reader.Xlsx
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
            private readonly bool[] _styleIsDate;
            private int[] _sharedOffsets;
            private readonly Utf8StringCache? _contentCache;
            private readonly ZipArchiveEntry? _entry;
            private int _nextCol;

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

            internal Enumerator(XlsxReader reader, ZipArchiveEntry entry, CancellationToken ct)
                : base(reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes), WorkbookLookups.InitialBufferCapacity(entry.Length), ct)
            {
                _reader = reader;
                _styleIsDate = reader._styleIsDate;
                _sharedOffsets = reader._sharedOffsets;
                _contentCache = reader._options.InternStrings ? new Utf8StringCache() : null;
                _entry = entry;
            }

            private protected override Stream OpenSource()
            {
                _reader.EnsureSharedLoaded();
                _sharedOffsets = _reader._sharedOffsets;
                return WorkbookLookups.OpenEntryStream(_entry!, _reader._decompressedBytes, _reader._options);
            }

            private protected override async ValueTask<Stream> OpenSourceAsync()
            {
                await _reader.EnsureSharedLoadedAsync(_ct).ConfigureAwait(false);
                _sharedOffsets = _reader._sharedOffsets;
                return await WorkbookLookups.OpenEntryStreamAsync(_entry!, _reader._decompressedBytes, _reader._options, _ct).ConfigureAwait(false);
            }

            /// <inheritdoc/>
            public Row Current =>
                new(_acc.CellSpan, _acc.ValueSpan, _reader.SharedSpan, _buf.AsSpan(0, _len), _reader.SharedStringCache, _contentCache);


            /// <inheritdoc/>
            public bool MoveNext()
            {
                if (!_nsChecked)
                {
                    DetectNamespace();
                }
                while (true)
                {
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
                                ParseRowBody();
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

            /// <inheritdoc/>
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
                                break;
                            }
                            return skipResult.Value;
                    }
                }
            }

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

                int rowStart = _pos;
                if (ParseRowInWindow())
                {
                    return new ValueTask<bool>(true);
                }
                if (_eof)
                {
                    ParseTruncatedRow(rowStart);
                    return new ValueTask<bool>(true);
                }
                return ParseRowBodySlowAsync(rowStart);
            }

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
                    return null;
                }
                return new ValueTask<bool>(false);
            }

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

            private void ParseRowBody()
            {
                int rowStart = _pos;
                while (!ParseRowInWindow())
                {
                    if (_eof)
                    {
                        ParseTruncatedRow(rowStart);
                        return;
                    }
                    RestartRowAt(rowStart);
                    Fill();
                    rowStart = _pos;
                }
            }

            private async ValueTask<bool> ParseRowBodySlowAsync(int rowStart)
            {
                while (true)
                {
                    RestartRowAt(rowStart);
                    await FillAsync().ConfigureAwait(false);
                    rowStart = _pos;
                    if (ParseRowInWindow())
                    {
                        return true;
                    }
                    if (_eof)
                    {
                        ParseTruncatedRow(rowStart);
                        return true;
                    }
                }
            }

            private void RestartRowAt(int rowStart)
            {
                _pos = rowStart;
                _acc.Reset();
                _nextCol = 0;
            }

            private void ParseTruncatedRow(int rowStart)
            {
                RestartRowAt(rowStart);
                ParseRow(_len);
            }

            private void DetectNamespace()
            {
                _nsChecked = true;
                Ensure(256);
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

            private bool BeginRow()
            {
                int gt = IndexOf((byte)'>');
                if (gt < 0)
                {
                    return MissingRowOpenTag();
                }
                return BeginRowAt(gt);
            }
            private bool ParseRowInWindow()
            {
                byte[] buf = _buf;
                int len = _len;
                int p = _pos;
                ReadOnlySpan<byte> rowEnd = _ns is null ? "</row"u8 : _ns.RowEnd;
                while (true)
                {
                    int lt = p < len && buf[p] == (byte)'<' ? p : IndexOfBounded(buf, len, p, (byte)'<');
                    if (lt < 0)
                    {
                        return false;
                    }
                    p = lt;
                    if (IsCellStart(buf, len, p))
                    {
                        p = ParseCellSpan(buf, len, p);
                        continue;
                    }
                    if (buf.AsSpan(p, Math.Min(rowEnd.Length, len - p)).StartsWith(rowEnd))
                    {
                        int gt = IndexOfBounded(buf, len, p, (byte)'>');
                        if (gt < 0)
                        {
                            return false;
                        }
                        _pos = gt + 1;
                        return true;
                    }
                    if (!SkipMarkupSpan(buf, len, ref p))
                    {
                        return false;
                    }
                }
            }

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
                        break;
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

            private int ParseCellSpan(byte[] buf, int len, int p)
            {
                if (!TryScanCanonicalCellTag(buf, len, p, out int gt, out int col, out int style, out Kind kind))
                {
                    gt = IndexOfBounded(buf, len, p, (byte)'>');
                    if (gt < 0)
                    {
                        return len;
                    }
                    ScanCellTagGeneric(buf.AsSpan(p, gt - p + 1), out col, out style, out kind);
                }
                if (col < 0)
                {
                    col = _nextCol;
                }
                _nextCol = col + 1;
                p = gt + 1;
                if (buf[gt - 1] == (byte)'/')
                {
                    return p;
                }

                if (buf.AsSpan(p, Math.Min(3, len - p)).StartsWith("<v>"u8))
                {
                    int valueStart = p + 3;
                    if (kind == Kind.Shared)
                    {
                        int end = TryEmitSharedIndex(buf, len, valueStart, col, style);
                        if (end >= 0)
                        {
                            return end;
                        }
                    }
                    int lt = IndexOfBounded(buf, len, valueStart, (byte)'<');
                    if (lt >= 0 && buf.AsSpan(lt, Math.Min(8, len - lt)).StartsWith("</v></c>"u8))
                    {
                        ReadOnlySpan<byte> value = buf.AsSpan(valueStart, lt - valueStart);
                        if (kind == Kind.Shared)
                        {
                            EmitShared(value, col, style);
                        }
                        else
                        {
                            EmitScalarValueFast(kind, value, valueStart, col, style);
                        }
                        return lt + 8;
                    }
                }
                else if (kind == Kind.Inline)
                {
                    int end = TryEmitPlainInline(buf, len, p, col, style);
                    if (end >= 0)
                    {
                        return end;
                    }
                }

                ReadOnlySpan<byte> cClose = CClose;
                int cEnd = IndexOfSeqBounded(buf, len, p, cClose);
                if (cEnd < 0)
                {
                    return len;
                }
                EmitCell(kind, buf.AsSpan(p, cEnd - p), col, style);
                return cEnd + cClose.Length;
            }

            private int TryEmitSharedIndex(byte[] buf, int len, int valueStart, int col, int style)
            {
                int i = valueStart;
                int index = 0;
                while (i < len && (uint)(buf[i] - '0') <= 9)
                {
                    index = (index * 10) + (buf[i] - '0');
                    i++;
                }
                if (i - valueStart is 0 or > 9 || !buf.AsSpan(i, Math.Min(8, len - i)).StartsWith("</v></c>"u8))
                {
                    return -1;
                }
                var (start, length, sharedIndex) = WorkbookLookups.SharedAt(_sharedOffsets, index);
                _acc.Add(col, start, length, CellType.ExcelString, style, CellValueSource.Shared, sharedIndex: sharedIndex);
                return i + 8;
            }

            private int TryEmitPlainInline(byte[] buf, int len, int p, int col, int style)
            {
                if (!buf.AsSpan(p, Math.Min(7, len - p)).StartsWith("<is><t>"u8))
                {
                    return -1;
                }
                int textStart = p + 7;
                int rel = buf.AsSpan(textStart, len - textStart).IndexOfAny((byte)'<', (byte)'&');
                if (rel < 0)
                {
                    return -1;
                }
                int lt = textStart + rel;
                if (!buf.AsSpan(lt, Math.Min(13, len - lt)).StartsWith("</t></is></c>"u8))
                {
                    return -1;
                }
                _acc.Add(col, textStart, rel, CellType.ExcelString, style, CellValueSource.RowBuffer);
                return lt + 13;
            }

            private enum HeadKind { End, Row, Skip }

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

            private static bool StartsWithElement(ReadOnlySpan<byte> span, ReadOnlySpan<byte> token)
            {
                return span.StartsWith(token) && (span.Length == token.Length || IsBoundary(span[token.Length]));
            }

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
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

            // Accepts only ` r="A1"`, ` s="12"`, ` t="s"` (any order, one space, double quotes); anything else goes generic.
            private static bool TryScanCanonicalCellTag(byte[] buf, int len, int i, out int gt, out int col, out int style, out Kind kind)
            {
                gt = -1;
                col = -1;
                style = 0;
                kind = Kind.Number;
                i += 2;
                while (true)
                {
                    if (len - i < 6)
                    {
                        return false;
                    }
                    byte b = buf[i];
                    if (b == (byte)'>')
                    {
                        gt = i;
                        return true;
                    }
                    if (b == (byte)'/')
                    {
                        gt = i + 1;
                        return buf[gt] == (byte)'>';
                    }
                    if (b != (byte)' ' || buf[i + 2] != (byte)'=' || buf[i + 3] != (byte)'"')
                    {
                        return false;
                    }
                    byte name = buf[i + 1];
                    i += 4;
                    int valueStart = i;
                    switch (name)
                    {
                        case (byte)'r':
                            int letters = 0;
                            while (i < len && (uint)(buf[i] - 'A') <= 25)
                            {
                                letters = (letters * 26) + (buf[i] - 'A' + 1);
                                i++;
                            }
                            if (i - valueStart is 0 or > 3)
                            {
                                return false;
                            }
                            col = letters - 1;
                            while (i < len && (uint)(buf[i] - '0') <= 9)
                            {
                                i++;
                            }
                            break;
                        case (byte)'s':
                            int value = 0;
                            while (i < len && (uint)(buf[i] - '0') <= 9)
                            {
                                value = (value * 10) + (buf[i] - '0');
                                i++;
                            }
                            if (i - valueStart is 0 or > 9)
                            {
                                return false;
                            }
                            style = value;
                            break;
                        case (byte)'t':
                            while (i < len && buf[i] != (byte)'"')
                            {
                                i++;
                            }
                            kind = ClassifyKind(buf.AsSpan(valueStart, i - valueStart));
                            break;
                        default:
                            return false;
                    }
                    if (i >= len || buf[i] != (byte)'"')
                    {
                        return false;
                    }
                    i++;
                }
            }

            [MethodImpl(MethodImplOptions.NoInlining)]
            private static void ScanCellTagGeneric(ReadOnlySpan<byte> open, out int col, out int style, out Kind kind)
            {
                ScanCellAttributes(open, out var rRef, out var sVal, out var tVal);
                col = XlsxXml.ColumnIndex(rRef);
                style = XlsxXml.ParseIntOr(sVal, 0);
                kind = ClassifyKind(tVal);
            }

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

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private static bool IsXmlSpace(byte b)
            {
                return b is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n';
            }

            private enum Kind { Number, Shared, Inline, Bool, Error, Formula, IsoDate }

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
                    3 => Kind.Formula,
                    9 => Kind.Inline,
                    _ => Kind.Number,
                };
            }

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

            private void EmitScalarValue(Kind kind, ReadOnlySpan<byte> v, int col, int style)
            {
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
                if (kind == Kind.Formula)
                {
                    AppendDecoded(v);
                    _acc.Add(col, vStart, _acc.ValueLength - vStart, cellType, style, CellValueSource.RowValues);
                    return;
                }
                AppendRaw(v);
                double number = 0;
                bool hasNumber = kind == Kind.Number && FastDouble.TryParse(v, out number);
                _acc.Add(col, vStart, _acc.ValueLength - vStart, cellType, style, CellValueSource.RowValues,
                    number: number, hasNumber: hasNumber);
            }

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
                    double serial = dt.ToOADate();
                    int start = _acc.ValueLength;
                    Span<byte> dst = _acc.ReserveValueSpan(32);
                    CellFormatter.TryFormatDouble(serial, dst, out int written);
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
                if (FastDate.TryParse(utf8, out value))
                {
                    return value.Year >= 100 || RejectDate(out value);
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
                return value.Year >= 100 || RejectDate(out value);
            }

            private static bool RejectDate(out DateTime value)
            {
                value = default;
                return false;
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

            private enum MarkupSeq { CommentEnd, CDataEnd }

            private int FindSeq(MarkupSeq seq, int start)
            {
                int rel = seq == MarkupSeq.CommentEnd
                    ? _buf.AsSpan(start, _len - start).IndexOf("-->"u8)
                    : _buf.AsSpan(start, _len - start).IndexOf("]]>"u8);
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
