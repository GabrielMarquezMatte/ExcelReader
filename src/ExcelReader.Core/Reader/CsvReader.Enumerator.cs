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
        public sealed class Enumerator : PooledStreamRowEnumerator, IExcelRowEnumerator
        {
            private const byte Cr = (byte)'\r';
            private const byte Lf = (byte)'\n';

            // Borrowed: CsvReader owns the stream's lifetime (it may be reused across enumerations).
            [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed", Justification = "Borrowed, not owned.")]
            private readonly byte _delimiter;
            private readonly byte _quote;
            private readonly bool _stripBom;

            private bool _bomChecked;

            private int _col;

            // Current field's bytes, built incrementally as either a single contiguous run in _buf
            // (the common case — zero-copy) or, once a discontiguous append is needed (a doubled ""
            // quote, or bytes trailing a closing quote), materialized into _acc's value buffer instead.
            // Scoped to a single TryParseRecordFromBuffer call and threaded by ref, alongside the local
            // `buf`/`len`/`pos` cursor, so the hot per-field loop touches only locals/registers instead
            // of chasing the _io indirection (BufferedStreamCursor) and instance-field loads per field.
            private struct FieldState
            {
                public int BufStart;
                public int BufLen;
                public bool Materialized;
                public int MatStart;
            }

            internal Enumerator(Stream stream, CsvReaderOptions options, CancellationToken ct = default)
                : base(stream, options.MaxCellBytes, nameof(CsvReaderOptions.MaxCellBytes), 64 * 1024, ct)
            {
                _delimiter = options.Delimiter;
                _quote = options.Quote;
                _stripBom = options.DetectEncodingFromByteOrderMark;
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

            // Fast-path gate mirrors MoveNextAsync's: once BOM-checked, a failed TryParseRecordFromBuffer
            // only ever happens with `_pos < _len` already true (it's restored to `start`, a position
            // that was `< _len` when the attempt began), and Fill only grows `_len` or sets Eof — so once
            // the loop is entered, `_pos < _len` is a loop invariant and never needs re-checking; only the
            // very first record (or a buffer genuinely exhausted between records) needs the slow prologue.
            public bool MoveNext()
            {
                if (!_bomChecked || _pos >= _len)
                {
                    EnsureBomStripped();
                    Ensure(1);
                    if (_pos >= _len)
                    {
                        return false;
                    }
                }
                while (true)
                {
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

            // Non-async fast path: once the BOM check is done, ~99.9% of records are already fully
            // buffered, so try the buffer-only parse synchronously before paying for an async state
            // machine at all. Only a genuine buffer miss (or the first-ever call, for the BOM check)
            // falls to the slow awaiting path.
            public ValueTask<bool> MoveNextAsync()
            {
                if (!_bomChecked || _pos >= _len)
                {
                    return MoveNextSlowAsync();
                }
                BeginRecord();
                int start = _pos;
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
                byte[] buf = _buf;
                int len = _len;
                byte delim = _delimiter;
                byte quote = _quote;
                int pos = _pos;

                // Fast path: a single long-span IndexOfAny finds whichever of {CR, LF, quote} comes
                // first — typically 50-200+ bytes, wide enough to actually engage the vectorized path
                // (per-field spans below are 5-20 bytes and mostly don't). If a terminator is hit before
                // any quote, the line up to it is provably quote-free, so it's split by delimiter
                // directly with no FieldState/materialization machinery at all. Falls through to the
                // general per-field parser when a quote is hit first, or when the terminator/EOF isn't
                // resolvable yet.
                ReadOnlySpan<byte> remaining = buf.AsSpan(pos, len - pos);
                int first = remaining.IndexOfAny(Cr, Lf, quote);
                bool quoteFirst = first >= 0 && remaining[first] == quote;
                int lineTerm = quoteFirst ? -1 : first;
                bool ambiguousCr = lineTerm >= 0 && remaining[lineTerm] == Cr && lineTerm == remaining.Length - 1 && !_eof;
                if (!quoteFirst && !ambiguousCr && (lineTerm >= 0 || _eof))
                {
                    ReadOnlySpan<byte> line = lineTerm >= 0 ? remaining[..lineTerm] : remaining;
                    ParseUnquotedLine(line, pos);
                    pos += line.Length;
                    if (lineTerm >= 0)
                    {
                        pos += buf[pos] == Cr && pos + 1 < len && buf[pos + 1] == Lf ? 2 : 1;
                    }
                    _pos = pos;
                    return true;
                }

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
                    int term = TryScanUnquotedRun(buf, len, delim, ref pos, ref f);
                    if (term == NeedMore)
                    {
                        _pos = pos;
                        return false;
                    }
                    CommitField(f);
                    if (term == RecordEnd)
                    {
                        _pos = pos;
                        return true;
                    }
                }
            }

            // Splits an entire quote-free line (already known fully buffered) by delimiter in one tight
            // loop — no FieldState, no materialization check, no quote branch per field, since none of
            // that machinery can possibly be needed here. `lineStart` is the line's absolute offset in
            // buf, so committed cells still point directly into it (zero-copy, same as the general path).
            private void ParseUnquotedLine(ReadOnlySpan<byte> line, int lineStart)
            {
                byte delim = _delimiter;
                int start = 0;
                while (true)
                {
                    int rel = line[start..].IndexOf(delim);
                    if (rel < 0)
                    {
                        int fieldLen = line.Length - start;
                        _acc.Add(_col++, lineStart + start, fieldLen, fieldLen == 0 ? CellType.Empty : CellType.ExcelString, style: 0, fromShared: false);
                        return;
                    }
                    _acc.Add(_col++, lineStart + start, rel, rel == 0 ? CellType.Empty : CellType.ExcelString, style: 0, fromShared: false);
                    start += rel + 1;
                }
            }

            // `pos` is right after the opening quote. Appends unescaped content ("" -> ") to _acc
            // and leaves `pos` right after the closing quote (or at EOF for an unterminated field).
            // False means the closing quote (or the byte after it, needed to rule out "") is not
            // buffered yet.
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
                        return true; // unterminated quoted field at EOF
                    }
                    int q = pos + rel;
                    if (q + 1 >= len && !_eof)
                    {
                        return false; // can't distinguish a closing quote from "" yet
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

            // Scans from `pos` (either the field's start, or right after a closing quote) for the
            // next delimiter/terminator, appending the run verbatim.
            private int TryScanUnquotedRun(ReadOnlySpan<byte> buf, int len, byte delim, ref int pos, ref FieldState f)
            {
                int rel = pos < len ? buf[pos..len].IndexOfAny(delim, Cr, Lf) : -1;
                if (rel < 0)
                {
                    if (!_eof)
                    {
                        return NeedMore;
                    }
                    FieldAppendBufRun(buf, pos, len - pos, ref f);
                    pos = len;
                    return RecordEnd;
                }
                int found = pos + rel;
                byte b = buf[found];
                if (b == Cr && found + 1 >= len && !_eof)
                {
                    return NeedMore; // can't tell a bare CR from CRLF yet
                }
                FieldAppendBufRun(buf, pos, found - pos, ref f);
                if (b == delim)
                {
                    pos = found + 1;
                    return FieldEnd;
                }
                pos = found + (b == Cr && found + 1 < len && buf[found + 1] == Lf ? 2 : 1);
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

            // Appends a run of bytes already sitting at buf[start..start+len). Stays a zero-copy
            // slice of buf as long as each run continues exactly where the previous one ended;
            // any gap (or a prior literal byte append) forces materialization into _acc from then on.
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

            // Appends one literal byte (the unescaped '"' from a doubled ""), which can never be a
            // contiguous continuation of the surrounding buf run.
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
                    _acc.Add(_col++, f.BufStart, f.BufLen, f.BufLen == 0 ? CellType.Empty : CellType.ExcelString, style: 0, fromShared: false);
                    return;
                }
                int len = _acc.ValueLength - f.MatStart;
                _acc.Add(_col++, f.MatStart, len, len == 0 ? CellType.Empty : CellType.ExcelString, style: 0, fromShared: true);
            }

            // --- buffer management (shared with XlsxReader/XlsbReader via BufferedStreamCursor) ---

            public void Dispose()
            {
                ReturnBuffers();
            }

            public ValueTask DisposeAsync()
            {
                ReturnBuffers();
                return ValueTask.CompletedTask;
            }

        }
    }
}
