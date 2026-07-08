using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
{
    public sealed partial class CsvReader
    {
        // Forward-only record scanner over a pooled buffer, structured like XlsxReader.Enumerator:
        // a single moving cursor (_pos) into _buf, refilled/compacted via BufferedStreamCursor, with
        // every field's bytes (unescaped as needed) copied into CellAccumulator's persistent value
        // buffer so that Current stays valid across the buffer compaction that the next MoveNext may trigger.
        // Records are parsed from buffered bytes only (TryParseRecordFromBuffer, no stream I/O)
        // when a record is only partially buffered the parse restarts after a refill, so sync and
        // async share one parser and the async path awaits once per refill, not per field.
        [SuppressMessage("Design", "CA1034:Nested types should not be visible",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public sealed class Enumerator : IExcelRowEnumerator
        {
            private const byte Cr = (byte)'\r';
            private const byte Lf = (byte)'\n';

            // Borrowed: CsvReader owns the stream's lifetime (it may be reused across enumerations).
            [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed", Justification = "Borrowed, not owned.")]
            [SuppressMessage("SharpSource", "SS066:Disposable field is not disposed", Justification = "Borrowed, not owned.")]
            private readonly Stream _stream;
            private readonly CancellationToken _ct;
            private readonly byte _delimiter;
            private readonly byte _quote;
            private readonly bool _stripBom;

            private readonly BufferedStreamCursor _io;
            private byte[] _buf => _io.Buf;
            private int _pos { get => _io.Pos; set => _io.Pos = value; }
            private int _len => _io.Len;
            private bool _eof => _io.Eof;
            private bool _bomChecked;

            private readonly CellAccumulator _acc; // per-record decoded values + cell descriptors
            private int _col;

            internal Enumerator(Stream stream, CsvReaderOptions options, CancellationToken ct = default)
            {
                _stream = stream;
                _ct = ct;
                _delimiter = options.Delimiter;
                _quote = options.Quote;
                _stripBom = options.DetectEncodingFromByteOrderMark;
                _io = new BufferedStreamCursor(options.MaxCellBytes, nameof(CsvReaderOptions.MaxCellBytes));
                _acc = new CellAccumulator(options.MaxCellBytes, nameof(CsvReaderOptions.MaxCellBytes));
            }

            public Row Current => new(_acc.CellSpan, _acc.ValueSpan, default);

            // Dense field access for CsvEnumerable<T>: CSV cells are stored contiguously in column
            // order (no gaps), so field i is _acc.CellSpan[i] — O(1), skipping Row's binary search and
            // the RowCells re-walk the generic projector would do.
            internal int FieldCount => _acc.Count;

            internal Cell FieldAt(int index)
            {
                CellDesc d = _acc.CellSpan[index];
                return new Cell(d.Type, _acc.ValueSpan.Slice(d.Start, d.Length));
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
                _acc.Reset();
                _col = 0;
            }

            // --- record/field parsing (buffer-only, no stream I/O) ---

            // Outcomes of TryScanUnquotedRun.
            private const int NeedMore = 0;
            private const int FieldEnd = 1;
            private const int RecordEnd = 2;

            // Parses one full record from _buf[_pos.._len]. Returns false when the record is not
            // fully buffered yet (and not EOF); the caller restores _pos to the record start,
            // refills, and re-parses the record from scratch (BeginRecord resets _acc).
            // ponytail: restart-on-refill re-scans the partial record after every Fill — fine while
            // records are far smaller than the 64KB buffer; make the parse resumable if huge records
            // over trickling streams ever matter.
            private bool TryParseRecordFromBuffer()
            {
                while (true)
                {
                    int valStart = _acc.ValueLength;
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
                    AddCell(valStart, _acc.ValueLength - valStart);
                    if (term == RecordEnd)
                    {
                        return true;
                    }
                }
            }

            // _pos is right after the opening quote. Appends unescaped content ("" -> ") to _acc
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

            // --- value/cell accumulation (CellAccumulator, shared with the XLSX/XLSB/XLS enumerators) ---

            private void AppendToVals(ReadOnlySpan<byte> src)
            {
                if (src.IsEmpty)
                {
                    return;
                }
                Span<byte> dst = _acc.ReserveValueSpan(src.Length);
                src.CopyTo(dst);
                _acc.Advance(src.Length);
            }

            private void AppendByteToVals(byte b)
            {
                _acc.AppendByte(b);
            }

            private void AddCell(int start, int len)
            {
                _acc.Add(_col++, start, len, len == 0 ? CellType.Empty : CellType.ExcelString, style: 0, fromShared: false);
            }

            // --- buffer management (shared with XlsxReader/XlsbReader via BufferedStreamCursor) ---

            private void Fill()
            {
                _io.Fill(_stream);
            }

            private ValueTask FillAsync()
            {
                return _io.FillAsync(_stream, _ct);
            }

            private void Ensure(int n)
            {
                _io.Ensure(_stream, n);
            }

            private ValueTask EnsureAsync(int n)
            {
                return _io.EnsureAsync(_stream, n, _ct);
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
                _io.Return();
                _acc.Return();
            }
        }
    }
}
