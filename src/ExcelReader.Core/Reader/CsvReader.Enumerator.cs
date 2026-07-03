using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
{
    public sealed partial class CsvReader
    {
        // Forward-only record scanner over a pooled buffer, structured like XlsxReader.Enumerator:
        // a single moving cursor (_pos) into _buf, refilled/compacted by Fill/PrepareBuffer, with
        // every field's bytes (unescaped as needed) copied into a persistent _vals buffer so that
        // Current stays valid across the buffer compaction that the next MoveNext may trigger.
        // Records are parsed from buffered bytes only (TryParseRecordFromBuffer, no stream I/O)
        // when a record is only partially buffered the parse restarts after a refill, so sync and
        // async share one parser and the async path awaits once per refill, not per field.
        [SuppressMessage("Design", "CA1034:Nested types should not be visible",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public sealed class Enumerator : IExcelRowEnumerator
        {
            private const int InitialBuf = 64 * 1024;
            private const int InitialVals = 4 * 1024;
            private const int InitialCells = 32;
            private const byte Cr = (byte)'\r';
            private const byte Lf = (byte)'\n';

            // Borrowed: CsvReader owns the stream's lifetime (it may be reused across enumerations).
            [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed", Justification = "Borrowed, not owned.")]
            [SuppressMessage("SharpSource", "SS066:Disposable field is not disposed", Justification = "Borrowed, not owned.")]
            private readonly Stream _stream;
            private readonly CancellationToken _ct;
            private readonly byte _delimiter;
            private readonly byte _quote;
            private readonly int _maxCellBytes;
            private readonly bool _stripBom;

            private byte[] _buf;
            private int _pos;
            private int _len;
            private bool _eof;
            private bool _bomChecked;

            private byte[] _vals;    // decoded (unescaped) bytes for every field of the current record
            private int _valLen;
            private CellDesc[] _cells;
            private int _col;

            internal Enumerator(Stream stream, CsvReaderOptions options, CancellationToken ct = default)
            {
                _stream = stream;
                _ct = ct;
                _delimiter = options.Delimiter;
                _quote = options.Quote;
                _maxCellBytes = options.MaxCellBytes;
                _stripBom = options.DetectEncodingFromByteOrderMark;
                _buf = ArrayPool<byte>.Shared.Rent(InitialBuf);
                _vals = ArrayPool<byte>.Shared.Rent(InitialVals);
                _cells = ArrayPool<CellDesc>.Shared.Rent(InitialCells);
            }

            public Row Current => new(_cells.AsSpan(0, FieldCount), _vals.AsSpan(0, _valLen), default);

            // Dense field access for CsvEnumerable<T>: CSV cells are stored contiguously in column
            // order (no gaps), so field i is _cells[i] — O(1), skipping Row's binary search and the
            // RowCells re-walk the generic projector would do.
            internal int FieldCount { get; private set; }

            internal Cell FieldAt(int index)
            {
                ref readonly CellDesc d = ref _cells[index];
                return new Cell(d.Type, _vals.AsSpan(d.Start, d.Length));
            }

            public bool MoveNext()
            {
                EnsureBomStripped();
                while (true)
                {
                    Ensure(1);
                    if (_pos >= _len)
                    {
                        return false;
                    }
                    BeginRecord();
                    int start = _pos;
                    if (TryParseRecordFromBuffer())
                    {
                        return true;
                    }
                    _pos = start;
                    Fill();
                }
            }

            public async ValueTask<bool> MoveNextAsync()
            {
                await EnsureBomStrippedAsync().ConfigureAwait(false);
                while (true)
                {
                    await EnsureAsync(1).ConfigureAwait(false);
                    if (_pos >= _len)
                    {
                        return false;
                    }
                    BeginRecord();
                    int start = _pos;
                    if (TryParseRecordFromBuffer())
                    {
                        return true;
                    }
                    _pos = start;
                    await FillAsync().ConfigureAwait(false);
                }
            }

            private void BeginRecord()
            {
                FieldCount = 0;
                _valLen = 0;
                _col = 0;
            }

            // --- record/field parsing (buffer-only, no stream I/O) ---

            // Outcomes of TryScanUnquotedRun.
            private const int NeedMore = 0;
            private const int FieldEnd = 1;
            private const int RecordEnd = 2;

            // Parses one full record from _buf[_pos.._len]. Returns false when the record is not
            // fully buffered yet (and not EOF); the caller restores _pos to the record start,
            // refills, and re-parses the record from scratch (BeginRecord resets _vals/_cells).
            // ponytail: restart-on-refill re-scans the partial record after every Fill — fine while
            // records are far smaller than the 64KB buffer; make the parse resumable if huge records
            // over trickling streams ever matter.
            private bool TryParseRecordFromBuffer()
            {
                while (true)
                {
                    int valStart = _valLen;
                    if (_pos < _len && _buf[_pos] == _quote)
                    {
                        _pos++;
                        if (!TryParseQuotedContent())
                        {
                            return false;
                        }
                    }
                    int term = TryScanUnquotedRun();
                    if (term == NeedMore)
                    {
                        return false;
                    }
                    AddCell(valStart, _valLen - valStart);
                    if (term == RecordEnd)
                    {
                        return true;
                    }
                }
            }

            // _pos is right after the opening quote. Appends unescaped content ("" -> ") to _vals
            // and leaves _pos right after the closing quote (or at EOF for an unterminated field).
            // False means the closing quote (or the byte after it, needed to rule out "") is not
            // buffered yet.
            private bool TryParseQuotedContent()
            {
                while (true)
                {
                    int rel = _pos < _len ? _buf.AsSpan(_pos, _len - _pos).IndexOf(_quote) : -1;
                    if (rel < 0)
                    {
                        if (!_eof)
                        {
                            return false;
                        }
                        AppendToVals(_buf.AsSpan(_pos, _len - _pos));
                        _pos = _len;
                        return true; // unterminated quoted field at EOF
                    }
                    int q = _pos + rel;
                    if (q + 1 >= _len && !_eof)
                    {
                        return false; // can't distinguish a closing quote from "" yet
                    }
                    AppendToVals(_buf.AsSpan(_pos, q - _pos));
                    if (q + 1 < _len && _buf[q + 1] == _quote)
                    {
                        AppendByteToVals(_quote);
                        _pos = q + 2;
                        continue;
                    }
                    _pos = q + 1;
                    return true;
                }
            }

            // Scans from _pos (either the field's start, or right after a closing quote) for the
            // next delimiter/terminator, appending the run verbatim.
            private int TryScanUnquotedRun()
            {
                int rel = _pos < _len ? _buf.AsSpan(_pos, _len - _pos).IndexOfAny(_delimiter, Cr, Lf) : -1;
                if (rel < 0)
                {
                    if (!_eof)
                    {
                        return NeedMore;
                    }
                    AppendToVals(_buf.AsSpan(_pos, _len - _pos));
                    _pos = _len;
                    return RecordEnd;
                }
                int found = _pos + rel;
                byte b = _buf[found];
                if (b == Cr && found + 1 >= _len && !_eof)
                {
                    return NeedMore; // can't tell a bare CR from CRLF yet
                }
                AppendToVals(_buf.AsSpan(_pos, found - _pos));
                if (b == _delimiter)
                {
                    _pos = found + 1;
                    return FieldEnd;
                }
                _pos = found + (b == Cr && found + 1 < _len && _buf[found + 1] == Lf ? 2 : 1);
                return RecordEnd;
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
                Ensure(3);
                if (_len - _pos >= 3 && _buf[_pos] == 0xEF && _buf[_pos + 1] == 0xBB && _buf[_pos + 2] == 0xBF)
                {
                    _pos += 3;
                }
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
                await EnsureAsync(3).ConfigureAwait(false);
                if (_len - _pos >= 3 && _buf[_pos] == 0xEF && _buf[_pos + 1] == 0xBB && _buf[_pos + 2] == 0xBF)
                {
                    _pos += 3;
                }
            }

            // --- vals / cells buffers ---

            private void AppendToVals(ReadOnlySpan<byte> src)
            {
                if (src.IsEmpty)
                {
                    return;
                }
                EnsureValsCapacity(_valLen + src.Length);
                src.CopyTo(_vals.AsSpan(_valLen));
                _valLen += src.Length;
            }

            private void AppendByteToVals(byte b)
            {
                EnsureValsCapacity(_valLen + 1);
                _vals[_valLen++] = b;
            }

            private void EnsureValsCapacity(int needed)
            {
                if (needed <= _vals.Length)
                {
                    return;
                }
                byte[] bigger = ArrayPool<byte>.Shared.Rent(
                    LimitChecks.NextBufferSize(_maxCellBytes, nameof(CsvReaderOptions.MaxCellBytes), _vals.Length, needed));
                Array.Copy(_vals, bigger, _valLen);
                ArrayPool<byte>.Shared.Return(_vals);
                _vals = bigger;
            }

            private void AddCell(int start, int len)
            {
                if (FieldCount == _cells.Length)
                {
                    CellDesc[] bigger = ArrayPool<CellDesc>.Shared.Rent(_cells.Length * 2);
                    Array.Copy(_cells, bigger, FieldCount);
                    ArrayPool<CellDesc>.Shared.Return(_cells);
                    _cells = bigger;
                }
                _cells[FieldCount++] = new CellDesc
                {
                    Column = _col++,
                    Start = start,
                    Length = len,
                    Type = len == 0 ? CellType.Empty : CellType.ExcelString,
                    Style = 0,
                    FromShared = false,
                };
            }

            // --- buffer management (mirrors XlsxReader.Enumerator) ---

            private void PrepareBuffer()
            {
                if (_pos > 0)
                {
                    _buf.AsSpan(_pos, _len - _pos).CopyTo(_buf);
                    _len -= _pos;
                    _pos = 0;
                }
                else if (_len == _buf.Length)
                {
                    byte[] bigger = ArrayPool<byte>.Shared.Rent(
                        LimitChecks.NextBufferSize(_maxCellBytes, nameof(CsvReaderOptions.MaxCellBytes), _buf.Length, _buf.Length + 1));
                    _buf.AsSpan(0, _len).CopyTo(bigger);
                    ArrayPool<byte>.Shared.Return(_buf);
                    _buf = bigger;
                }
            }

            private void Fill()
            {
                PrepareBuffer();
                int n = _stream.Read(_buf, _len, _buf.Length - _len);
                if (n == 0)
                {
                    _eof = true;
                }
                else
                {
                    _len += n;
                }
            }

            private async ValueTask FillAsync()
            {
                PrepareBuffer();
                int n = await _stream.ReadAsync(_buf.AsMemory(_len, _buf.Length - _len), _ct).ConfigureAwait(false);
                if (n == 0)
                {
                    _eof = true;
                }
                else
                {
                    _len += n;
                }
            }

            private void Ensure(int n)
            {
                while (_len - _pos < n && !_eof)
                {
                    Fill();
                }
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

            public void Dispose()
            {
                ReturnBuffers();
            }

            public ValueTask DisposeAsync()
            {
                ReturnBuffers();
                return ValueTask.CompletedTask;
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
