using System.Buffers;
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
        public sealed class Enumerator : IDisposable, IAsyncDisposable
        {
            private const int InitialBuf = 64 * 1024;
            private const int InitialVals = 4 * 1024;
            private const int InitialCells = 32;

            // Borrowed: the reader outlives the enumerator and owns its own disposal — do not dispose here.
            [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed", Justification = "Borrowed, not owned.")]
            [SuppressMessage("SharpSource", "SS066:Disposable field is not disposed", Justification = "Borrowed, not owned.")]
            private readonly XlsxReader _reader;
            private readonly CancellationToken _ct; // honored only by the async path
            // Owned: opened by Get(Async)Enumerator for this enumerator alone; disposed in Dispose(Async).
            [SuppressMessage("SharpSource", "SS066:Disposable field is not disposed", Justification = "Disposed in Dispose().")]
            private Stream? _sheet;
            private byte[] _buf;
            private int _pos;
            private int _len;
            private bool _eof;

            private byte[] _vals;     // decoded values for the current row (numbers/inline/formula text)
            private int _valLen;
            private CellDesc[] _cells; // current row's cells, ascending by column
            private int _cellCount;
            private int _nextCol;

            internal Enumerator(XlsxReader reader, Stream sheet, CancellationToken ct = default)
            {
                _reader = reader;
                _sheet = sheet;
                _ct = ct;
                _buf = ArrayPool<byte>.Shared.Rent(InitialBuf);
                _vals = ArrayPool<byte>.Shared.Rent(InitialVals);
                _cells = ArrayPool<CellDesc>.Shared.Rent(InitialCells);
            }

            public Row Current =>
                new(_cells.AsSpan(0, _cellCount), _vals.AsSpan(0, _valLen), _reader.SharedSpan);

            // The sync and async row scanners share every span-touching helper below (ClassifyHead,
            // ClassifyRowHead, ReadCellOpenTag, EmitCell). The two families differ only in whether the
            // buffer refill (Fill / FillAsync) and the byte searches await — so the span work is factored
            // into sync helpers that never hold a span across an await.

            public bool MoveNext()
            {
                while (true)
                {
                    int lt = IndexOf((byte)'<');
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
                            BeginRow(out bool selfClose);
                            if (!selfClose)
                            {
                                ParseRow();
                            }
                            return true;
                        default:
                            int skip = IndexOf((byte)'>');
                            if (skip < 0)
                            {
                                return false;
                            }
                            _pos = skip + 1;
                            break;
                    }
                }
            }

            public async ValueTask<bool> MoveNextAsync()
            {
                while (true)
                {
                    int lt = await IndexOfAsync((byte)'<').ConfigureAwait(false);
                    if (lt < 0)
                    {
                        return false;
                    }
                    _pos = lt;
                    await EnsureAsync(12).ConfigureAwait(false);
                    switch (ClassifyHead())
                    {
                        case HeadKind.End:
                            return false;
                        case HeadKind.Row:
                            BeginRow(out bool selfClose);
                            if (!selfClose)
                            {
                                await ParseRowAsync().ConfigureAwait(false);
                            }
                            return true;
                        default:
                            int skip = await IndexOfAsync((byte)'>').ConfigureAwait(false);
                            if (skip < 0)
                            {
                                return false;
                            }
                            _pos = skip + 1;
                            break;
                    }
                }
            }

            // Consumes the <row ...> open tag and resets per-row state. Call only after ClassifyHead()==Row.
            private void BeginRow(out bool selfClose)
            {
                int gt = IndexOf((byte)'>'); // open tag already fully buffered by the Ensure(12) above
                selfClose = _buf[gt - 1] == '/';
                _pos = gt + 1;
                _cellCount = 0;
                _valLen = 0;
                _nextCol = 0;
            }

            private void ParseRow()
            {
                while (true)
                {
                    int lt = IndexOf((byte)'<');
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
                            int skip = IndexOf((byte)'>');
                            if (skip < 0)
                            {
                                return;
                            }
                            _pos = skip + 1;
                            break;
                    }
                }
            }

            private async ValueTask ParseRowAsync()
            {
                while (true)
                {
                    int lt = await IndexOfAsync((byte)'<').ConfigureAwait(false);
                    if (lt < 0)
                    {
                        return;
                    }
                    _pos = lt;
                    await EnsureAsync(6).ConfigureAwait(false);
                    switch (ClassifyRowHead())
                    {
                        case RowHead.EndRow:
                            int gt = await IndexOfAsync((byte)'>').ConfigureAwait(false);
                            _pos = gt < 0 ? _len : gt + 1;
                            return;
                        case RowHead.Cell:
                            await ParseCellAsync().ConfigureAwait(false);
                            break;
                        default:
                            int skip = await IndexOfAsync((byte)'>').ConfigureAwait(false);
                            if (skip < 0)
                            {
                                return;
                            }
                            _pos = skip + 1;
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
                int cEnd = IndexOfSeq("</c>"u8); // ensures whole cell contiguous; shifts _pos consistently
                if (cEnd < 0)
                {
                    _pos = _len;
                    return;
                }
                EmitCell(header.Kind, _buf.AsSpan(_pos, cEnd - _pos), header.Col, header.Style);
                _pos = cEnd + 4;
            }

            private async ValueTask ParseCellAsync()
            {
                int gt = await IndexOfAsync((byte)'>').ConfigureAwait(false);
                var header = ReadCellOpenTag(gt); // sync: no span survives into the next await
                if (header.SelfClose)
                {
                    return;
                }
                int cEnd = await IndexOfCloseCellAsync().ConfigureAwait(false);
                if (cEnd < 0)
                {
                    _pos = _len;
                    return;
                }
                EmitCell(header.Kind, _buf.AsSpan(_pos, cEnd - _pos), header.Col, header.Style);
                _pos = cEnd + 4;
            }

            private enum HeadKind { End, Row, Skip }

            private HeadKind ClassifyHead()
            {
                var head = _buf.AsSpan(_pos, Math.Min(12, _len - _pos));
                if (head.StartsWith("</sheetData"u8) || head.StartsWith("</worksheet"u8))
                {
                    return HeadKind.End;
                }
                if (head.StartsWith("<row"u8) && (head.Length < 5 || IsBoundary(head[4])))
                {
                    return HeadKind.Row;
                }
                return HeadKind.Skip;
            }

            private enum RowHead { EndRow, Cell, Other }

            private RowHead ClassifyRowHead()
            {
                var head = _buf.AsSpan(_pos, Math.Min(6, _len - _pos));
                if (head.StartsWith("</row"u8))
                {
                    return RowHead.EndRow;
                }
                if (head.StartsWith("<c"u8) && (head.Length < 3 || IsBoundary(head[2])))
                {
                    return RowHead.Cell;
                }
                return RowHead.Other;
            }

            private readonly record struct CellHeader(int Col, int Style, Kind Kind, bool SelfClose);

            // Parses the <c ...> open tag ending at `gt`, advances _pos past it, and returns the extracted
            // (non-span) attributes. Holds spans only within this synchronous call, so callers may await after.
            private CellHeader ReadCellOpenTag(int gt)
            {
                var open = _buf.AsSpan(_pos, gt - _pos + 1);
                var rRef = XlsxXml.Attr(open, " r=\""u8);
                var sVal = XlsxXml.Attr(open, " s=\""u8);
                var tVal = XlsxXml.Attr(open, " t=\""u8);

                int col = rRef.IsEmpty ? _nextCol : XlsxXml.ColumnIndex(rRef);
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

            private enum Kind { Number, Shared, Inline, Bool, Error, Formula }

            // "" / "n" -> Number; "s" shared; "inlineStr" inline; "b" bool; "e" error; "str" formula result.
            private static Kind ClassifyKind(ReadOnlySpan<byte> t)
            {
                if (t.SequenceEqual("s"u8))
                {
                    return Kind.Shared;
                }
                if (t.SequenceEqual("inlineStr"u8))
                {
                    return Kind.Inline;
                }
                if (t.SequenceEqual("b"u8))
                {
                    return Kind.Bool;
                }
                if (t.SequenceEqual("e"u8))
                {
                    return Kind.Error;
                }
                if (t.SequenceEqual("str"u8))
                {
                    return Kind.Formula;
                }

                return Kind.Number;
            }


            private void EmitCell(Kind kind, ReadOnlySpan<byte> inner, int col, int style)
            {
                // Shared strings: <v> holds an index; point the cell at that slice of the shared buffer.
                if (kind == Kind.Shared)
                {
                    var (start, len) = _reader.SharedAt(ParseInt(ElementText(inner, "<v>"u8, "</v>"u8)));
                    AddCell(col, start, len, CellType.ExcelString, style, fromShared: true);
                    return;
                }

                int vStart = _valLen;
                if (kind == Kind.Inline)
                {
                    EnsureValsCapacity(_valLen + inner.Length);
                    _valLen += XlsxXml.WriteTextRuns(inner, _vals.AsSpan(_valLen));
                    AddCell(col, vStart, _valLen - vStart, CellType.ExcelString, style, fromShared: false);
                    return;
                }

                CellType cellType = kind switch
                {
                    Kind.Bool => CellType.Boolean,
                    Kind.Error => CellType.Error,
                    Kind.Formula => CellType.Formula,
                    _ => _reader.IsDateStyle(style) ? CellType.Date : CellType.Number,
                };
                AppendDecoded(ElementText(inner, "<v>"u8, "</v>"u8));
                AddCell(col, vStart, _valLen - vStart, cellType, style, fromShared: false);
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
                EnsureValsCapacity(_valLen + src.Length);
                _valLen += XlsxXml.Decode(src, _vals.AsSpan(_valLen));
            }

            private void AddCell(int col, int start, int len, CellType type, int style, bool fromShared)
            {
                if (_cellCount == _cells.Length)
                {
                    var bigger = ArrayPool<CellDesc>.Shared.Rent(_cells.Length * 2);
                    Array.Copy(_cells, bigger, _cellCount);
                    ArrayPool<CellDesc>.Shared.Return(_cells);
                    _cells = bigger;
                }
                _cells[_cellCount++] = new CellDesc
                {
                    Column = col,
                    Start = start,
                    Length = len,
                    Type = type,
                    Style = style,
                    FromShared = fromShared,
                };
            }

            private void EnsureValsCapacity(int needed)
            {
                if (needed <= _vals.Length)
                {
                    return;
                }
                var bigger = ArrayPool<byte>.Shared.Rent(Math.Max(_vals.Length * 2, needed));
                Array.Copy(_vals, bigger, _valLen);
                ArrayPool<byte>.Shared.Return(_vals);
                _vals = bigger;
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

            private void Ensure(int n)
            {
                while (_len - _pos < n && !_eof)
                {
                    Fill();
                }
            }

            // Make room for more bytes (compact consumed prefix, else grow), then read once.
            private void Fill()
            {
                if (_pos > 0)
                {
                    _buf.AsSpan(_pos, _len - _pos).CopyTo(_buf);
                    _len -= _pos;
                    _pos = 0;
                }
                else if (_len == _buf.Length)
                {
                    var bigger = ArrayPool<byte>.Shared.Rent(_buf.Length * 2);
                    _buf.AsSpan(0, _len).CopyTo(bigger);
                    ArrayPool<byte>.Shared.Return(_buf);
                    _buf = bigger;
                }
                int n = _sheet!.Read(_buf, _len, _buf.Length - _len);
                if (n == 0)
                {
                    _eof = true;
                }
                else
                {
                    _len += n;
                }
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

            // Only one multi-byte terminator is searched on the async path, so it's hardcoded rather than
            // taking a ReadOnlySpan<byte> (which can't be an async method parameter).
            private ValueTask<int> IndexOfCloseCellAsync()
            {
                int rel = _buf.AsSpan(_pos, _len - _pos).IndexOf("</c>"u8);
                if (rel >= 0)
                {
                    return new ValueTask<int>(_pos + rel);
                }
                return _eof ? new ValueTask<int>(-1) : IndexOfCloseCellSlowAsync();
            }

            private async ValueTask<int> IndexOfCloseCellSlowAsync()
            {
                do
                {
                    await FillAsync().ConfigureAwait(false);
                    int rel = _buf.AsSpan(_pos, _len - _pos).IndexOf("</c>"u8);
                    if (rel >= 0)
                    {
                        return _pos + rel;
                    }
                }
                while (!_eof);
                return -1;
            }

            private ValueTask EnsureAsync(int n)
            {
                if (_len - _pos >= n || _eof)
                {
                    return ValueTask.CompletedTask;
                }
                return EnsureSlowAsync(n);
            }

            private async ValueTask EnsureSlowAsync(int n)
            {
                while (_len - _pos < n && !_eof)
                {
                    await FillAsync().ConfigureAwait(false);
                }
            }

            private async ValueTask FillAsync()
            {
                if (_pos > 0)
                {
                    _buf.AsSpan(_pos, _len - _pos).CopyTo(_buf);
                    _len -= _pos;
                    _pos = 0;
                }
                else if (_len == _buf.Length)
                {
                    var bigger = ArrayPool<byte>.Shared.Rent(_buf.Length * 2);
                    _buf.AsSpan(0, _len).CopyTo(bigger);
                    ArrayPool<byte>.Shared.Return(_buf);
                    _buf = bigger;
                }
                int n = await _sheet!.ReadAsync(_buf.AsMemory(_len, _buf.Length - _len), _ct).ConfigureAwait(false);
                if (n == 0)
                {
                    _eof = true;
                }
                else
                {
                    _len += n;
                }
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
                if (_buf.Length > 0)
                {
                    ArrayPool<byte>.Shared.Return(_buf);
                    _buf = [];
                }
                if (_vals.Length > 0)
                {
                    ArrayPool<byte>.Shared.Return(_vals);
                    _vals = [];
                }
                if (_cells.Length > 0)
                {
                    ArrayPool<CellDesc>.Shared.Return(_cells);
                    _cells = [];
                }
            }
        }
    }
}
