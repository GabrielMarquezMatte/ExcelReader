using System.Diagnostics.CodeAnalysis;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Reader.Csv
{
    internal enum FieldScanOutcome
    {
        NeedMore,
        FieldEnd,
        RecordEnd,
    }

    internal enum SimpleRecordOutcome
    {
        Done,
        NeedMore,
        Generic,
    }

    public sealed partial class CsvReader
    {
        /// <summary>A forward-only cursor over a <see cref="CsvReader"/> source's records, reading either synchronously or asynchronously.</summary>
        /// <remarks>
        /// A single moving cursor (<c>_pos</c>) into a pooled buffer refilled/compacted via
        /// <see cref="BufferedStreamCursor"/>. Most fields reference the buffer directly with no copy;
        /// only fields needing unescaping materialize into the cell accumulator's value buffer.
        /// Records parse from buffered bytes only — a partially-buffered record restarts after a
        /// refill, so sync and async share one parser and async awaits once per refill, not per field.
        /// </remarks>
        [SuppressMessage("Design", "CA1034:Nested types should not be visible",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public sealed class Enumerator : PooledStreamRowEnumerator, IExcelRowEnumerator
        {
            private const byte Cr = (byte)'\r';
            private const byte Lf = (byte)'\n';
            private const int BufferExhausted = -1;
            private const int NeedsGeneric = -2;
            private const int BatchRecords = 128;
            private const int BatchCells = 256;

            private readonly byte _delimiter;
            private readonly byte _quote;
            private readonly bool _stripBom;

            private bool _bomChecked;

            private int _col;

            private long _recordStart;

            private CsvStructuralScanner _scanner;
            private bool _scannerValid;
            private readonly bool _scannerSupported;

            private readonly int[] _batchCellEnds = new int[BatchRecords];
            private readonly int[] _batchRecordStarts = new int[BatchRecords];
            private int _batchSize;
            private int _batchNext;
            private int _firstCell;
            private int _cellCount;

            private readonly Utf8StringCache? _contentCache;

            private struct FieldState
            {
                public int BufStart;
                public int BufLen;
                public bool Materialized;
                public int MatStart;
            }

            internal Enumerator(Stream stream, CsvReaderOptions options, CancellationToken ct = default)
                : base(stream, options.MaxCellBytes, nameof(CsvReaderOptions.MaxCellBytes), 64 * 1024, ownsSource: false, ct)
            {
                _delimiter = options.Delimiter;
                _quote = options.Quote;
                _stripBom = options.DetectEncodingFromByteOrderMark;
                _contentCache = options.InternStrings ? new Utf8StringCache() : null;
                _scanner = new CsvStructuralScanner(_delimiter, _quote);
                _scannerSupported = CsvStructuralScanner.Supports(_delimiter, _quote);
            }

            internal Enumerator(ReadOnlyMemory<byte> content, CsvReaderOptions options, CancellationToken ct = default)
                : base(content, options.MaxCellBytes, nameof(CsvReaderOptions.MaxCellBytes), ct)
            {
                _delimiter = options.Delimiter;
                _quote = options.Quote;
                _stripBom = options.DetectEncodingFromByteOrderMark;
                _contentCache = options.InternStrings ? new Utf8StringCache() : null;
                _scanner = new CsvStructuralScanner(_delimiter, _quote);
                _scannerSupported = CsvStructuralScanner.Supports(_delimiter, _quote);
            }

            /// <inheritdoc/>
            public Row Current => new(RecordCells, _buf.AsSpan(0, _len), _acc.ValueSpan, rowBuffer: default, sharedStringCache: null, contentCache: _contentCache);

            internal long CurrentRecordStart => _recordStart;

            internal int FieldCount => _cellCount;

            private ReadOnlySpan<CellDesc> RecordCells
            {
                get
                {
                    return _acc.RawCells.AsSpan(_firstCell, _cellCount);
                }
            }

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            internal Cell FieldAt(int index)
            {
                ref readonly CellDesc d = ref RecordCells[index];
                return d.ToCell(_buf.AsSpan(0, _len), _acc.ValueSpan, rowBuffer: default, sharedStringCache: null, contentCache: _contentCache);
            }

            /// <inheritdoc/>
            public bool MoveNext()
            {
                if (_batchNext < _batchSize)
                {
                    DeliverBatched();
                    return true;
                }
                if (!_bomChecked || _pos >= _len)
                {
                    EnsureBomStripped();
                    EnsureInvalidatingScanner(1);
                    if (_pos >= _len)
                    {
                        return false;
                    }
                }
                while (true)
                {
                    BeginRecord();
                    int start = _pos;
                    _recordStart = _io.BaseOffset + start;
                    if (TryParseRecordFromBuffer())
                    {
                        return true;
                    }
                    _pos = start;
                    FillInvalidatingScanner();
                }
            }

            /// <inheritdoc/>
            public ValueTask<bool> MoveNextAsync()
            {
                if (_batchNext < _batchSize)
                {
                    DeliverBatched();
                    return new ValueTask<bool>(true);
                }
                if (!_bomChecked || _pos >= _len)
                {
                    return MoveNextSlowAsync();
                }
                BeginRecord();
                int start = _pos;
                _recordStart = _io.BaseOffset + start;
                if (TryParseRecordFromBuffer())
                {
                    return new ValueTask<bool>(true);
                }
                _pos = start;
                return MoveNextSlowAsync();
            }

            private async ValueTask<bool> MoveNextSlowAsync()
            {
                await EnsureBomStrippedAsync().ConfigureAwait(false);
                while (true)
                {
                    await EnsureInvalidatingScannerAsync(1).ConfigureAwait(false);
                    if (_pos >= _len)
                    {
                        return false;
                    }
                    BeginRecord();
                    int start = _pos;
                    _recordStart = _io.BaseOffset + start;
                    if (TryParseRecordFromBuffer())
                    {
                        return true;
                    }
                    _pos = start;
                    await FillInvalidatingScannerAsync().ConfigureAwait(false);
                }
            }


            private void EnsureInvalidatingScanner(int count)
            {
                _scannerValid = false;
                Ensure(count);
            }

            private ValueTask EnsureInvalidatingScannerAsync(int count)
            {
                _scannerValid = false;
                return EnsureAsync(count);
            }

            private void FillInvalidatingScanner()
            {
                _scannerValid = false;
                Fill();
            }

            private ValueTask FillInvalidatingScannerAsync()
            {
                _scannerValid = false;
                return FillAsync();
            }

            private void BeginRecord()
            {
                _acc.Reset();
                _col = 0;
                _batchSize = 0;
                _batchNext = 0;
            }

            private void DeliverBatched()
            {
                int record = _batchNext++;
                _firstCell = record == 0 ? 0 : _firstCell + _cellCount;
                _cellCount = _batchCellEnds[record] - _firstCell;
                _recordStart = _io.BaseOffset + _batchRecordStarts[record];
            }

            // ponytail: restart-on-refill re-scans the partial record after every Fill — fine while
            private bool TryParseRecordFromBuffer()
            {
                int len = _len;
                int recordStart = _pos;

                SimpleRecordOutcome simple = TryParseSimpleRecord(len, recordStart, out int genericFieldStart);
                if (simple == SimpleRecordOutcome.Done)
                {
                    DeliverBatched();
                    return true;
                }
                _scannerValid = false;
                if (simple == SimpleRecordOutcome.NeedMore)
                {
                    _pos = recordStart;
                    return false;
                }
                _firstCell = 0;
                if (!TryParseGenericRest(len, genericFieldStart))
                {
                    return false;
                }
                _cellCount = _acc.Count;
                return true;
            }

            [MethodImpl(MethodImplOptions.NoInlining)]
            private bool TryParseGenericRest(int len, int pos)
            {
                byte delim = _delimiter;
                byte quote = _quote;
                ReadOnlySpan<byte> buf = _buf.AsSpan(0, len);
                FieldState f = default;

                while (true)
                {
                    f.BufStart = pos;
                    f.BufLen = 0;
                    f.Materialized = false;

                    if (pos < len && buf[pos] == quote)
                    {
                        pos++;
                        if (!TryParseQuotedContent(buf, len, quote, ref pos, ref f))
                        {
                            _pos = pos;
                            return false;
                        }
                    }
                    FieldScanOutcome term = TryScanUnquotedRun(buf, len, delim, ref pos, ref f);
                    if (term == FieldScanOutcome.NeedMore)
                    {
                        _pos = pos;
                        return false;
                    }
                    CommitField(f);
                    if (term == FieldScanOutcome.RecordEnd)
                    {
                        _pos = pos;
                        return true;
                    }
                }
            }

            private SimpleRecordOutcome TryParseSimpleRecord(int len, int pos, out int genericFieldStart)
            {
                genericFieldStart = pos;
                if (!_scannerSupported)
                {
                    return SimpleRecordOutcome.Generic;
                }
                if (!_scannerValid)
                {
                    _scanner.Reset(_buf, len, pos);
                    _scannerValid = true;
                }
                _acc.ReserveCells(BatchCells);
                int fieldStart = pos;
                _batchRecordStarts[0] = pos;
                int recordEnd = DrainFields(_buf, ref fieldStart);
                genericFieldStart = fieldStart;
                if (recordEnd >= 0)
                {
                    _pos = recordEnd;
                    return SimpleRecordOutcome.Done;
                }
                if (recordEnd == NeedsGeneric)
                {
                    return SimpleRecordOutcome.Generic;
                }
                if (!_eof)
                {
                    return SimpleRecordOutcome.NeedMore;
                }
                if (_scanner.EndsInsideQuotes)
                {
                    return SimpleRecordOutcome.Generic;
                }
                CellDesc last = DescribeField(_buf, fieldStart, len - fieldStart, _scanner.Escapes != 0, _col);
                _acc.Add(_col++, last.Start, last.Length, last.Type, style: 0, last.Source);
                _batchCellEnds[0] = _acc.Count;
                _batchRecordStarts[0] = pos;
                _batchSize = 1;
                _pos = len;
                return SimpleRecordOutcome.Done;
            }

            [MethodImpl(MethodImplOptions.NoInlining)]
            private int DrainFields(byte[] buf, ref int fieldStart)
            {
                ref CellDesc cells = ref MemoryMarshal.GetArrayDataReference(_acc.RawCells);
                int n = _acc.Count;
                int first = n;
                int start = fieldStart;
                ulong ends = 0;
                ulong escapes = 0;
                int blockStart = 0;
                int result;
                while (true)
                {
                    if (ends == 0)
                    {
                        ulong carried = escapes != 0 ? 1UL : 0UL;
                        if (!_scanner.TryTakeBlock())
                        {
                            result = _scanner.Blocked ? NeedsGeneric : BufferExhausted;
                            break;
                        }
                        ends = _scanner.Ends;
                        escapes = _scanner.Escapes | carried;
                        blockStart = _scanner.BlockStart;
                        if (!BlockFits(n, first, ends))
                        {
                            result = NeedsGeneric;
                            break;
                        }
                        continue;
                    }
                    int bit = BitOperations.TrailingZeroCount(ends);
                    int stop = blockStart + bit;
                    bool isRecordEnd = buf[stop] != _delimiter;
                    if (isRecordEnd && IsCrAwaitingLf(buf, stop))
                    {
                        result = BufferExhausted;
                        break;
                    }
                    int length = stop - start;
                    if (length != 0 && buf[start] == _quote)
                    {
                        ulong escapesInField = escapes & ((2UL << bit) - 1);
                        escapes ^= escapesInField;
                        Unsafe.Add(ref cells, n) = DescribeQuotedField(buf, start, length, escapesInField != 0, n - first);
                    }
                    else
                    {
                        Unsafe.Add(ref cells, n) = TextField(n - first, start, length, CellValueSource.RowValues);
                    }
                    n++;
                    ends &= ends - 1;
                    start = stop + 1;
                    if (isRecordEnd)
                    {
                        start = SkipLfAfterCr(buf, stop);
                        int recordCells = n - first;
                        first = n;
                        if (AddBatchedRecord(n, recordCells, start))
                        {
                            result = start;
                            break;
                        }
                    }
                }
                _scanner.PutBack(ends, escapes);
                return FinishDrain(n, first, start, result, ref fieldStart);
            }

            private bool BlockFits(int n, int first, ulong ends)
            {
                int fields = BitOperations.PopCount(ends);
                return n + fields <= _acc.RawCells.Length && n - first + fields <= ExcelLimits.MaxColumns;
            }

            private int FinishDrain(int n, int first, int start, int result, ref int fieldStart)
            {
                if (_batchSize > 0 && result < 0)
                {
                    n = first;
                    start = _batchRecordStarts[_batchSize];
                    result = start;
                    _scannerValid = false;
                }
                _acc.CommitAscending(n, n - first - 1);
                _col = n - first;
                fieldStart = start;
                return result;
            }

            private bool AddBatchedRecord(int cellEnd, int recordCells, int nextRecordStart)
            {
                int record = _batchSize++;
                _batchCellEnds[record] = cellEnd;
                if (_batchSize == BatchRecords || cellEnd + recordCells > BatchCells)
                {
                    return true;
                }
                _batchRecordStarts[_batchSize] = nextRecordStart;
                return false;
            }

            private bool IsCrAwaitingLf(byte[] buf, int recordTerminator)
            {
                return buf[recordTerminator] == Cr && recordTerminator + 1 >= _len && !_eof;
            }

            private int SkipLfAfterCr(byte[] buf, int recordTerminator)
            {
                int next = recordTerminator + 1;
                if (buf[recordTerminator] == Cr && next < _len && buf[next] == Lf)
                {
                    return next + 1;
                }
                return next;
            }

            private CellDesc DescribeField(byte[] buf, int start, int length, bool escaped, int col)
            {
                if (length != 0 && buf[start] == _quote)
                {
                    return DescribeQuotedField(buf, start, length, escaped, col);
                }
                return TextField(col, start, length, CellValueSource.RowValues);
            }

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private CellDesc DescribeQuotedField(byte[] buf, int start, int length, bool escaped, int col)
            {
                if (escaped)
                {
                    return EscapedField(buf.AsSpan(start + 1, length - 2), col);
                }
                return TextField(col, start + 1, length - 2, CellValueSource.RowValues);
            }

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private static CellDesc TextField(int col, int start, int length, CellValueSource source)
            {
                return new CellDesc
                {
                    Column = col,
                    Start = start,
                    Length = length,
                    Type = length == 0 ? CellType.Empty : CellType.ExcelString,
                    Source = source,
                    SharedIndex = -1,
                };
            }

            [MethodImpl(MethodImplOptions.NoInlining)]
            private CellDesc EscapedField(ReadOnlySpan<byte> content, int col)
            {
                int valueStart = _acc.ValueLength;
                Span<byte> dst = _acc.ReserveValueSpan(content.Length);
                int written = 0;
                while (true)
                {
                    int quoteAt = content.IndexOf(_quote);
                    if (quoteAt < 0)
                    {
                        content.CopyTo(dst[written..]);
                        written += content.Length;
                        break;
                    }
                    content[..(quoteAt + 1)].CopyTo(dst[written..]);
                    written += quoteAt + 1;
                    content = content[(quoteAt + 2)..];
                }
                _acc.Advance(written);
                return TextField(col, valueStart, written, CellValueSource.Shared);
            }

            private bool TryParseQuotedContent(ReadOnlySpan<byte> buf, int len, byte quote, ref int pos, ref FieldState f)
            {
                while (true)
                {
                    int rel = pos < len ? buf[pos..len].IndexOf(quote) : -1;
                    if (rel < 0)
                    {
                        if (!_eof)
                        {
                            return false;
                        }
                        FieldAppendBufRun(buf, pos, len - pos, ref f);
                        pos = len;
                        return true;
                    }
                    int q = pos + rel;
                    if (q + 1 >= len && !_eof)
                    {
                        return false;
                    }
                    FieldAppendBufRun(buf, pos, q - pos, ref f);
                    if (q + 1 < len && buf[q + 1] == quote)
                    {
                        FieldAppendLiteralByte(buf, quote, ref f);
                        pos = q + 2;
                        continue;
                    }
                    pos = q + 1;
                    return true;
                }
            }

            private FieldScanOutcome TryScanUnquotedRun(ReadOnlySpan<byte> buf, int len, byte delim, ref int pos, ref FieldState f)
            {
                int rel = pos < len ? buf[pos..len].IndexOfAny(delim, Cr, Lf) : -1;
                if (rel < 0)
                {
                    if (!_eof)
                    {
                        return FieldScanOutcome.NeedMore;
                    }
                    FieldAppendBufRun(buf, pos, len - pos, ref f);
                    pos = len;
                    return FieldScanOutcome.RecordEnd;
                }
                int found = pos + rel;
                byte b = buf[found];
                if (b == Cr && found + 1 >= len && !_eof)
                {
                    return FieldScanOutcome.NeedMore;
                }
                FieldAppendBufRun(buf, pos, found - pos, ref f);
                if (b == delim)
                {
                    pos = found + 1;
                    return FieldScanOutcome.FieldEnd;
                }
                pos = found + (b == Cr && found + 1 < len && buf[found + 1] == Lf ? 2 : 1);
                return FieldScanOutcome.RecordEnd;
            }

            private void StripBomFromBuffer()
            {
                ReadOnlySpan<byte> buf = _buf.AsSpan(0, _len);
                if (_len - _pos >= 3 && buf[_pos] == 0xEF && buf[_pos + 1] == 0xBB && buf[_pos + 2] == 0xBF)
                {
                    _pos += 3;
                }
            }

            private void EnsureBomStripped()
            {
                if (_bomChecked)
                {
                    return;
                }
                _bomChecked = true;
                if (!_stripBom)
                {
                    return;
                }
                EnsureInvalidatingScanner(3);
                StripBomFromBuffer();
            }

            private async ValueTask EnsureBomStrippedAsync()
            {
                if (_bomChecked)
                {
                    return;
                }
                _bomChecked = true;
                if (!_stripBom)
                {
                    return;
                }
                await EnsureInvalidatingScannerAsync(3).ConfigureAwait(false);
                StripBomFromBuffer();
            }

            private void FieldAppendBufRun(ReadOnlySpan<byte> buf, int start, int len, ref FieldState f)
            {
                if (len == 0)
                {
                    return;
                }
                if (!f.Materialized)
                {
                    if (f.BufLen == 0)
                    {
                        f.BufStart = start;
                        f.BufLen = len;
                        return;
                    }
                    if (start == f.BufStart + f.BufLen)
                    {
                        f.BufLen += len;
                        return;
                    }
                    Materialize(buf, ref f);
                }
                Span<byte> dst = _acc.ReserveValueSpan(len);
                buf.Slice(start, len).CopyTo(dst);
                _acc.Advance(len);
            }

            private void FieldAppendLiteralByte(ReadOnlySpan<byte> buf, byte b, ref FieldState f)
            {
                if (!f.Materialized)
                {
                    Materialize(buf, ref f);
                }
                _acc.AppendByte(b);
            }

            private void Materialize(ReadOnlySpan<byte> buf, ref FieldState f)
            {
                f.MatStart = _acc.ValueLength;
                if (f.BufLen > 0)
                {
                    Span<byte> dst = _acc.ReserveValueSpan(f.BufLen);
                    buf.Slice(f.BufStart, f.BufLen).CopyTo(dst);
                    _acc.Advance(f.BufLen);
                }
                f.Materialized = true;
            }

            private void CommitField(FieldState f)
            {
                if (!f.Materialized)
                {
                    _acc.Add(_col++, f.BufStart, f.BufLen, f.BufLen == 0 ? CellType.Empty : CellType.ExcelString, style: 0, CellValueSource.RowValues);
                    return;
                }
                int len = _acc.ValueLength - f.MatStart;
                _acc.Add(_col++, f.MatStart, len, len == 0 ? CellType.Empty : CellType.ExcelString, style: 0, CellValueSource.Shared);
            }

        }
    }
}
