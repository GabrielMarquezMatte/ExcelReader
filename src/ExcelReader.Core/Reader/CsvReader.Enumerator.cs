using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
{
    public sealed partial class CsvReader
    {
        // Forward-only record scanner over a pooled buffer, structured like XlsxReader.Enumerator:
        // a single moving cursor (_pos) into _buf, refilled/compacted via BufferedStreamCursor. Most
        // fields (unquoted, or quoted without a doubled "") are already contiguous bytes in _buf, so
        // cells reference that buffer directly with no copy; only fields needing unescaping fall back
        // to CellAccumulator's value buffer, whose contents stay valid across the compaction that the
        // next MoveNext may trigger. Records are parsed from buffered bytes only (TryParseRecordFromBuffer,
        // no stream I/O); when a record is only partially buffered the parse restarts after a refill, so
        // sync and async share one parser and the async path awaits once per refill, not per field.
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

            // Current field's bytes, built incrementally as either a single contiguous run in _buf
            // (the common case — zero-copy) or, once a discontiguous append is needed (a doubled ""
            // quote, or bytes trailing a closing quote), materialized into _acc's value buffer instead.
            private int _fBufStart;
            private int _fBufLen;
            private bool _fMaterialized;
            private int _fMatStart;

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

            // Cells point either into _buf (the common, zero-copy case: unquoted or plain-quoted
            // fields are already contiguous bytes read straight from the stream) or into _acc's value
            // buffer (only for fields needing unescaping, e.g. a doubled "" quote, or malformed
            // trailing bytes after a closing quote) — see CellDesc.FromShared / ToCell.
            public Row Current => new(_acc.CellSpan, _buf.AsSpan(0, _len), _acc.ValueSpan);

            // Dense field access for CsvEnumerable<T>: CSV cells are stored contiguously in column
            // order (no gaps), so field i is _acc.CellSpan[i] — O(1), skipping Row's binary search and
            // the RowCells re-walk the generic projector would do.
            internal int FieldCount => _acc.Count;

            internal Cell FieldAt(int index)
            {
                CellDesc d = _acc.CellSpan[index];
                ReadOnlySpan<byte> buf = d.FromShared ? _acc.ValueSpan : _buf.AsSpan(0, _len);
                return new Cell(d.Type, buf.Slice(d.Start, d.Length));
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
                    FieldBegin();
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
                    CommitField();
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
                        FieldAppendBufRun(_pos, _len - _pos);
                        _pos = _len;
                        return true; // unterminated quoted field at EOF
                    }
                    int q = _pos + rel;
                    if (q + 1 >= _len && !_eof)
                    {
                        return false; // can't distinguish a closing quote from "" yet
                    }
                    FieldAppendBufRun(_pos, q - _pos);
                    if (q + 1 < _len && _buf[q + 1] == _quote)
                    {
                        FieldAppendLiteralByte(_quote);
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
                    FieldAppendBufRun(_pos, _len - _pos);
                    _pos = _len;
                    return RecordEnd;
                }
                int found = _pos + rel;
                byte b = _buf[found];
                if (b == Cr && found + 1 >= _len && !_eof)
                {
                    return NeedMore; // can't tell a bare CR from CRLF yet
                }
                FieldAppendBufRun(_pos, found - _pos);
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

            // --- field building (zero-copy _buf slice, falling back to _acc's value buffer only
            // when a field's bytes aren't contiguous in _buf — a doubled "" quote or malformed bytes
            // trailing a closing quote) ---

            private void FieldBegin()
            {
                _fBufStart = _pos;
                _fBufLen = 0;
                _fMaterialized = false;
            }

            // Appends a run of bytes already sitting at _buf[start..start+len). Stays a zero-copy
            // slice of _buf as long as each run continues exactly where the previous one ended;
            // any gap (or a prior literal byte append) forces materialization into _acc from then on.
            private void FieldAppendBufRun(int start, int len)
            {
                if (len == 0)
                {
                    return;
                }
                if (!_fMaterialized)
                {
                    if (_fBufLen == 0)
                    {
                        _fBufStart = start;
                        _fBufLen = len;
                        return;
                    }
                    if (start == _fBufStart + _fBufLen)
                    {
                        _fBufLen += len;
                        return;
                    }
                    Materialize();
                }
                Span<byte> dst = _acc.ReserveValueSpan(len);
                _buf.AsSpan(start, len).CopyTo(dst);
                _acc.Advance(len);
            }

            // Appends one literal byte (the unescaped '"' from a doubled ""), which can never be a
            // contiguous continuation of the surrounding _buf run.
            private void FieldAppendLiteralByte(byte b)
            {
                if (!_fMaterialized)
                {
                    Materialize();
                }
                _acc.AppendByte(b);
            }

            private void Materialize()
            {
                _fMatStart = _acc.ValueLength;
                if (_fBufLen > 0)
                {
                    Span<byte> dst = _acc.ReserveValueSpan(_fBufLen);
                    _buf.AsSpan(_fBufStart, _fBufLen).CopyTo(dst);
                    _acc.Advance(_fBufLen);
                }
                _fMaterialized = true;
            }

            private void CommitField()
            {
                if (!_fMaterialized)
                {
                    _acc.Add(_col++, _fBufStart, _fBufLen, _fBufLen == 0 ? CellType.Empty : CellType.ExcelString, style: 0, fromShared: false);
                    return;
                }
                int len = _acc.ValueLength - _fMatStart;
                _acc.Add(_col++, _fMatStart, len, len == 0 ? CellType.Empty : CellType.ExcelString, style: 0, fromShared: true);
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
