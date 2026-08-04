using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
{
    public sealed partial class CsvReader
    {
        /// <summary>A forward-only cursor over a <see cref="CsvReader"/> source's records, reading either synchronously or asynchronously.</summary>
        /// <remarks>
        /// Structured like <c>XlsxReader.Enumerator</c>: a single moving cursor (<c>_pos</c>) into the pooled buffer,
        /// refilled/compacted via <see cref="BufferedStreamCursor"/>. Most fields (unquoted, or quoted without a doubled
        /// <c>""</c>) are already contiguous bytes in the buffer, so cells reference it directly with no copy; only
        /// fields needing unescaping fall back to the cell accumulator's value buffer, whose contents stay valid
        /// across the compaction that the next <c>MoveNext</c> may trigger. Records are parsed from buffered bytes
        /// only, with no stream I/O; when a record is only partially buffered the parse restarts after a refill, so
        /// the synchronous and asynchronous paths share one parser and the async path awaits once per refill, not
        /// per field. A quote-free record is emitted by <see cref="CsvControlScanner"/> in a single vectorized pass
        /// over its bytes; only a record containing a quote falls back to the per-field scalar parser below.
        /// </remarks>
        [SuppressMessage("Design", "CA1034:Nested types should not be visible",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public sealed class Enumerator : PooledStreamRowEnumerator, IExcelRowEnumerator
        {
            private const byte Cr = (byte)'\r';
            private const byte Lf = (byte)'\n';

            // Borrowed: CsvReader owns the stream's lifetime (it may be reused across enumerations).
            private readonly byte _delimiter;
            private readonly byte _quote;
            private readonly bool _stripBom;

            private bool _bomChecked;

            private int _col;

            // Content-keyed dedup cache for GetString(); see CsvReaderOptions.InternStrings. CSV has
            // no stable shared-string table index the way XLSX/XLSB/XLS do, so this is the only dedup
            // path available to it.
            private readonly Utf8StringCache? _contentCache;

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
                : base(stream, options.MaxCellBytes, nameof(CsvReaderOptions.MaxCellBytes), 64 * 1024, ownsSource: false, ct)
            {
                _delimiter = options.Delimiter;
                _quote = options.Quote;
                _stripBom = options.DetectEncodingFromByteOrderMark;
                _contentCache = options.InternStrings ? new Utf8StringCache() : null;
            }

            internal Enumerator(ReadOnlyMemory<byte> content, CsvReaderOptions options, CancellationToken ct = default)
                : base(content, options.MaxCellBytes, nameof(CsvReaderOptions.MaxCellBytes), ct)
            {
                _delimiter = options.Delimiter;
                _quote = options.Quote;
                _stripBom = options.DetectEncodingFromByteOrderMark;
                _contentCache = options.InternStrings ? new Utf8StringCache() : null;
            }

            // Cells point either into _buf (the common, zero-copy case: unquoted or plain-quoted
            // fields are already contiguous bytes read straight from the stream) or into _acc's value
            // buffer (only for fields needing unescaping, e.g. a doubled "" quote, or malformed
            // trailing bytes after a closing quote) — see CellDesc.Source / ToCell.
            /// <inheritdoc/>
            public Row Current => new(_acc.CellSpan, _buf.AsSpan(0, _len), _acc.ValueSpan, rowBuffer: default, sharedStringCache: null, contentCache: _contentCache);

            // Dense field access for CsvEnumerable<T>: CSV cells are stored contiguously in column
            // order (no gaps), so field i is _acc.CellSpan[i] — O(1), skipping Row's binary search and
            // the RowCells re-walk the generic projector would do.
            internal int FieldCount => _acc.Count;

            internal Cell FieldAt(int index)
            {
                ref readonly CellDesc d = ref _acc.CellSpan[index];
                return d.ToCell(_buf.AsSpan(0, _len), _acc.ValueSpan, rowBuffer: default, sharedStringCache: null, contentCache: _contentCache);
            }

            // Fast-path gate mirrors MoveNextAsync's: once BOM-checked, a failed TryParseRecordFromBuffer
            // only ever happens with `_pos < _len` already true (it's restored to `start`, a position
            // that was `< _len` when the attempt began), and Fill only grows `_len` or sets Eof — so once
            // the loop is entered, `_pos < _len` is a loop invariant and never needs re-checking; only the
            // very first record (or a buffer genuinely exhausted between records) needs the slow prologue.
            /// <inheritdoc/>
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
            /// <inheritdoc/>
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

            // Parses one full record from _buf[_pos.._len]. Returns false when the record is not
            // fully buffered yet (and not EOF); the caller restores _pos to the record start,
            // refills, and re-parses the record from scratch (BeginRecord resets _acc).
            // ponytail: restart-on-refill re-scans the partial record after every Fill — fine while
            // records are far smaller than the 64KB buffer; make the parse resumable if huge records
            // over trickling streams ever matter.
            private bool TryParseRecordFromBuffer()
            {
                ReadOnlySpan<byte> buf = _buf.AsSpan(0, _len);
                int len = _len;
                byte delim = _delimiter;
                byte quote = _quote;
                int recordStart = _pos;

                SimpleRecordOutcome simple = TryParseSimpleRecord(buf, len, recordStart);
                if (simple == SimpleRecordOutcome.Done)
                {
                    return true;
                }
                if (simple == SimpleRecordOutcome.NeedMore)
                {
                    _pos = recordStart;
                    return false;
                }

                // A quote turned up: the fields emitted above are discarded and the record is re-parsed
                // from its start by the general per-field path. Costs no more than before the fused fast
                // path existed — the old code also scanned for the quote and then re-parsed the record.
                _acc.Reset();
                _col = 0;
                int pos = recordStart;
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

            // Emits every field of a quote-free record in one vectorized pass: CsvControlScanner reports
            // delimiters and terminators in order from a single scan, so each byte is read once instead of
            // twice (a line-terminator search followed by a per-field delimiter search). Bails out with
            // SimpleQuoted the moment a quote is seen, since unescaping "" and honoring quoted
            // delimiters/newlines is the general path's job.
            // Measured (CsvReadBenchmark.ExcelReaderWide, 50000 rows x 32 columns, i7-1355U):
            // 11.134 ms -> 8.119 ms, ~27% faster. The narrow 4-column shape (ExcelReader) is flat
            // within noise (~4.0-4.4 ms either way) — expected, since mask reuse only pays off once
            // a record has enough fields to amortize a vector load.
            private SimpleRecordOutcome TryParseSimpleRecord(ReadOnlySpan<byte> buf, int len, int pos)
            {
                CsvControlScanner scanner = new(buf, pos, len, _delimiter, _quote);
                int fieldStart = pos;
                while (true)
                {
                    int stop = scanner.Next();
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
                        return SimpleRecordOutcome.NeedMore; // can't tell a bare CR from CRLF yet
                    }
                    AddField(fieldStart, stop - fieldStart);
                    _pos = stop + (b == Cr && stop + 1 < len && buf[stop + 1] == Lf ? 2 : 1);
                    return SimpleRecordOutcome.Done;
                }
            }

            private void AddField(int start, int length)
            {
                _acc.Add(_col++, start, length, length == 0 ? CellType.Empty : CellType.ExcelString,
                         style: 0, CellValueSource.RowValues);
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
                    return FieldScanOutcome.NeedMore; // can't tell a bare CR from CRLF yet
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

            // The three-byte UTF-8 BOM, if present at the very start. Splitting this out of the
            // sync/async entry points keeps the byte check itself in one place: only the buffer fill
            // differs between them.
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
                Ensure(3);
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
                await EnsureAsync(3).ConfigureAwait(false);
                StripBomFromBuffer();
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
                    _acc.Add(_col++, f.BufStart, f.BufLen, f.BufLen == 0 ? CellType.Empty : CellType.ExcelString, style: 0, CellValueSource.RowValues);
                    return;
                }
                int len = _acc.ValueLength - f.MatStart;
                _acc.Add(_col++, f.MatStart, len, len == 0 ? CellType.Empty : CellType.ExcelString, style: 0, CellValueSource.Shared);
            }

        }
    }
}
