using System.Diagnostics.CodeAnalysis;
using System.Numerics;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
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
        Quoted,
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
            private const int ChunkExhausted = -1;
            private const int NeedsGeneric = -2;

            private readonly byte _delimiter;
            private readonly byte _quote;
            private readonly bool _stripBom;

            private bool _bomChecked;

            private int _col;

            private long _recordStart;

            private CsvControlScanner _scanner;
            private bool _scannerValid;

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
                _scanner = new CsvControlScanner(_delimiter, _quote);
            }

            internal Enumerator(ReadOnlyMemory<byte> content, CsvReaderOptions options, CancellationToken ct = default)
                : base(content, options.MaxCellBytes, nameof(CsvReaderOptions.MaxCellBytes), ct)
            {
                _delimiter = options.Delimiter;
                _quote = options.Quote;
                _stripBom = options.DetectEncodingFromByteOrderMark;
                _contentCache = options.InternStrings ? new Utf8StringCache() : null;
                _scanner = new CsvControlScanner(_delimiter, _quote);
            }

            /// <inheritdoc/>
            public Row Current => new(_acc.CellSpan, _buf.AsSpan(0, _len), _acc.ValueSpan, rowBuffer: default, sharedStringCache: null, contentCache: _contentCache);

            internal long CurrentRecordStart => _recordStart;

            internal int FieldCount => _acc.Count;

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            internal Cell FieldAt(int index)
            {
                ref readonly CellDesc d = ref _acc.CellSpan[index];
                return d.ToCell(_buf.AsSpan(0, _len), _acc.ValueSpan, rowBuffer: default, sharedStringCache: null, contentCache: _contentCache);
            }

            /// <inheritdoc/>
            public bool MoveNext()
            {
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
            }

            // ponytail: restart-on-refill re-scans the partial record after every Fill — fine while
            private bool TryParseRecordFromBuffer()
            {
                int len = _len;
                byte delim = _delimiter;
                byte quote = _quote;
                int recordStart = _pos;

                SimpleRecordOutcome simple = TryParseSimpleRecord(len, recordStart, out int quotedFieldStart);
                if (simple == SimpleRecordOutcome.Done)
                {
                    return true;
                }
                _scannerValid = false;
                if (simple == SimpleRecordOutcome.NeedMore)
                {
                    _pos = recordStart;
                    return false;
                }

                ReadOnlySpan<byte> buf = _buf.AsSpan(0, len);
                int pos = quotedFieldStart;
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

            private SimpleRecordOutcome TryParseSimpleRecord(int len, int pos, out int quotedFieldStart)
            {
                quotedFieldStart = pos;
                if (_scannerValid)
                {
                    _scanner.Continue(_buf, len);
                }
                else
                {
                    _scanner.Reset(_buf, len, pos);
                    _scannerValid = true;
                }
                byte[] buf = _buf;
                int fieldStart = pos;
                while (true)
                {
                    int recordEnd = DrainFields(buf, ref fieldStart);
                    if (recordEnd >= 0)
                    {
                        _pos = recordEnd;
                        return SimpleRecordOutcome.Done;
                    }
                    int stop = _scanner.Next();
                    if (stop < 0)
                    {
                        if (!_eof)
                        {
                            return SimpleRecordOutcome.NeedMore;
                        }
                        AddField(fieldStart, len - fieldStart);
                        _pos = len;
                        return SimpleRecordOutcome.Done;
                    }
                    byte b = buf[stop];
                    if (b == _quote)
                    {
                        quotedFieldStart = fieldStart;
                        return SimpleRecordOutcome.Quoted;
                    }
                    if (b == _delimiter)
                    {
                        AddField(fieldStart, stop - fieldStart);
                        fieldStart = stop + 1;
                        continue;
                    }
                    if (b == Cr && stop + 1 >= len && !_eof)
                    {
                        return SimpleRecordOutcome.NeedMore;
                    }
                    AddField(fieldStart, stop - fieldStart);
                    bool isCrLf = b == Cr && stop + 1 < len && buf[stop + 1] == Lf;
                    if (isCrLf)
                    {
                        _scanner.SkipByte(stop + 1);
                    }
                    _pos = stop + (isCrLf ? 2 : 1);
                    return SimpleRecordOutcome.Done;
                }
            }

            private void AddField(int start, int length)
            {
                _acc.Add(_col++, start, length, length == 0 ? CellType.Empty : CellType.ExcelString,
                         style: 0, CellValueSource.RowValues);
            }

            [MethodImpl(MethodImplOptions.NoInlining)]
            private int DrainFields(byte[] buf, ref int fieldStart)
            {
                if (!_scanner.TryTakeMask(out uint mask, out int chunkStart))
                {
                    return NeedsGeneric;
                }
                CellAccumulator acc = _acc;
                CellDesc[] cells = acc.RawCells;
                byte delimiter = _delimiter;
                int n = acc.Count;
                int c = _col;
                int start = fieldStart;
                int recordEnd = ChunkExhausted;
                while (true)
                {
                    if (mask == 0)
                    {
                        if (_scanner.TryTakeMask(out mask, out chunkStart))
                        {
                            continue;
                        }
                        recordEnd = NeedsGeneric;
                        break;
                    }
                    int stop = chunkStart + BitOperations.TrailingZeroCount(mask);
                    byte b = buf[stop];
                    uint rest = mask & (mask - 1);
                    if (b == delimiter)
                    {
                        recordEnd = ChunkExhausted;
                    }
                    else if (b == Lf)
                    {
                        recordEnd = stop + 1;
                    }
                    else if (b == Cr && rest != 0 && chunkStart + BitOperations.TrailingZeroCount(rest) == stop + 1 && buf[stop + 1] == Lf)
                    {
                        recordEnd = stop + 2;
                        rest &= rest - 1;
                    }
                    else
                    {
                        recordEnd = NeedsGeneric;
                    }
                    if (recordEnd == NeedsGeneric || (uint)n >= (uint)cells.Length || (uint)c >= ExcelLimits.MaxColumns)
                    {
                        recordEnd = NeedsGeneric;
                        break;
                    }
                    int length = stop - start;
                    cells[n++] = new CellDesc
                    {
                        Column = c++,
                        Start = start,
                        Length = length,
                        Type = length == 0 ? CellType.Empty : CellType.ExcelString,
                        Source = CellValueSource.RowValues,
                        SharedIndex = -1,
                    };
                    start = stop + 1;
                    mask = rest;
                    if (recordEnd >= 0)
                    {
                        break;
                    }
                }
                acc.CommitAscending(n, c - 1);
                _col = c;
                fieldStart = start;
                _scanner.PutBack(mask, chunkStart);
                return recordEnd;
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
