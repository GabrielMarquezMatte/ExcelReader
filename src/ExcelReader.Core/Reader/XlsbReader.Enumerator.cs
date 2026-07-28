using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
{
    public sealed partial class XlsbReader
    {
        /// <summary>Forward-only enumerator over an <see cref="XlsbReader"/> sheet's rows.</summary>
        /// <remarks>
        /// Streams the underlying binary <c>sheetN.bin</c> entry through a refillable pooled buffer;
        /// <c>Biff12RecordReader</c> framing guarantees that a partial record at the buffer boundary
        /// is detected and retried after the next fill.
        /// </remarks>
        [SuppressMessage("Design", "CA1034:Nested types should not be visible",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public sealed class Enumerator : PooledStreamRowEnumerator, IExcelRowEnumerator
        {
            [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed",
                Justification = "XlsbReader is borrowed; its lifetime is managed by the caller, not this enumerator.")]
            private readonly XlsbReader _reader;
            private bool _ended;
            // A BrtRowHdr for the NEXT row was already consumed while collecting cells for the current row.
            // On the next MoveNext call, skip the "seek to row header" step.
            private bool _pendingRowHdr;

            internal Enumerator(XlsbReader reader, Stream sheet, long entryLength = 0, CancellationToken ct = default)
                : base(sheet, reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes), WorkbookLookups.InitialBufferCapacity(entryLength), ct)
            {
                _reader = reader;
            }

            /// <inheritdoc/>
            public Row Current =>
                new(_acc.CellSpan, _acc.ValueSpan, _reader.SharedSpan, rowBuffer: default, _reader.SharedStringCache);

            /// <inheritdoc/>
            public bool MoveNext()
            {
                _ct.ThrowIfCancellationRequested();
                return MoveNextCore();
            }

            // Non-async fast path mirroring MoveNextCore: with a 64KB buffer, almost every row's
            // header and cells are already fully buffered, so this runs the same *FromBuffer primitives
            // MoveNextCore uses, entirely synchronously, and only pays for an async state machine when
            // a primitive actually reports a buffer miss (result == 2) — a real refill, not per row.
            /// <inheritdoc/>
            public ValueTask<bool> MoveNextAsync()
            {
                _ct.ThrowIfCancellationRequested();
                if (_ended)
                {
                    return new ValueTask<bool>(false);
                }
                while (true)
                {
                    ResetRow();
                    if (!_pendingRowHdr)
                    {
                        int seek = SeekRowHdrFromBuffer();
                        if (seek == 0)
                        {
                            return new ValueTask<bool>(false);
                        }
                        if (seek == 2)
                        {
                            return MoveNextRowAsync(seekDone: false);
                        }
                    }
                    _pendingRowHdr = false;
                    int collect = CollectCellsFromBuffer();
                    if (collect == 2)
                    {
                        return MoveNextRowAsync(seekDone: true);
                    }
                    if (_acc.Count > 0)
                    {
                        return new ValueTask<bool>(true);
                    }
                    if (!_pendingRowHdr)
                    {
                        return new ValueTask<bool>(false);
                    }
                    // Empty row — loop again, still fully synchronous.
                }
            }

            // Resumes a row attempt that hit a buffer miss mid-seek (seekDone: false) or mid-collect
            // (seekDone: true — the seek step, and clearing _pendingRowHdr, already happened
            // synchronously in MoveNextAsync). _acc already holds whatever cells were collected before
            // the miss; CollectCellsAsync appends to it rather than restarting. Once this row is
            // resolved, continues the empty-row retry loop with the same awaiting primitives.
            private async ValueTask<bool> MoveNextRowAsync(bool seekDone)
            {
                if (!seekDone && !await SeekRowHdrAsync().ConfigureAwait(false))
                {
                    return false;
                }
                _pendingRowHdr = false;
                await CollectCellsAsync().ConfigureAwait(false);
                while (true)
                {
                    if (_acc.Count > 0)
                    {
                        return true;
                    }
                    if (!_pendingRowHdr)
                    {
                        return false;
                    }
                    ResetRow();
                    if (!await SeekRowHdrAsync().ConfigureAwait(false))
                    {
                        return false;
                    }
                    _pendingRowHdr = false;
                    await CollectCellsAsync().ConfigureAwait(false);
                }
            }

            private bool MoveNextCore()
            {
                if (_ended)
                {
                    return false;
                }
                while (true)
                {
                    ResetRow();
                    if (!_pendingRowHdr && !SkipToRowHdr())
                    {
                        return false;
                    }
                    _pendingRowHdr = false;
                    CollectCells();
                    if (_acc.Count > 0)
                    {
                        return true;
                    }
                    if (!_pendingRowHdr)
                    {
                        return false; // EOF or BrtEndSheetData with no cells in sight
                    }
                    // Empty row — advance to next
                }
            }

            // --- Sync record-loop helpers (use blocking Fill) ---

            private bool SkipToRowHdr()
            {
                while (TryNextRecord(out int id, out _))
                {
                    if (id == Brt.RowHdr)
                    {
                        return true;
                    }
                    if (IsEndSheetData(id))
                    {
                        _ended = true;
                        return false;
                    }
                }
                ThrowIfTruncated(); // TryNextRecord returned false => EOF (not EndSheetData)
                _ended = true;
                return false;
            }

            private void CollectCells()
            {
                while (TryNextRecord(out int id, out ReadOnlySpan<byte> payload))
                {
                    if (id == Brt.RowHdr)
                    {
                        _pendingRowHdr = true;
                        return;
                    }
                    if (IsEndSheetData(id))
                    {
                        _ended = true;
                        return;
                    }
                    ProcessCell(id, payload);
                }
                ThrowIfTruncated(); // EOF reached without an EndSheetData record
                _ended = true;
            }

            // --- Async record-loop helpers ---
            // Span-touching work stays in non-async helpers (*FromBuffer); the async wrappers only call
            // FillAsync and re-invoke those helpers — no ReadOnlySpan<byte> survives across an await.

            private async ValueTask<bool> SeekRowHdrAsync()
            {
                while (true)
                {
                    int result = SeekRowHdrFromBuffer();
                    if (result == 1)
                    {
                        return true;
                    }
                    if (result == 0)
                    {
                        return false;
                    }
                    await FillAsync().ConfigureAwait(false);
                }
            }

            private async ValueTask CollectCellsAsync()
            {
                while (true)
                {
                    int result = CollectCellsFromBuffer();
                    if (result != 2)
                    {
                        return;
                    }
                    await FillAsync().ConfigureAwait(false);
                }
            }

            // 1 = found BrtRowHdr; 0 = ended (EndSheetData/EOF); 2 = buffer exhausted (needs refill).
            private int SeekRowHdrFromBuffer()
            {
                while (TryNextRecordFromBuffer(out int id, out _))
                {
                    if (id == Brt.RowHdr)
                    {
                        return 1;
                    }
                    if (IsEndSheetData(id))
                    {
                        _ended = true;
                        return 0;
                    }
                }
                if (_eof)
                {
                    ThrowIfTruncated();
                    _ended = true;
                    return 0;
                }
                return 2;
            }

            // 1 = found next BrtRowHdr (_pendingRowHdr set); 0 = ended; 2 = needs refill.
            private int CollectCellsFromBuffer()
            {
                while (TryNextRecordFromBuffer(out int id, out ReadOnlySpan<byte> payload))
                {
                    if (id == Brt.RowHdr)
                    {
                        _pendingRowHdr = true;
                        return 1;
                    }
                    if (IsEndSheetData(id))
                    {
                        _ended = true;
                        return 0;
                    }
                    ProcessCell(id, payload);
                }
                if (_eof)
                {
                    ThrowIfTruncated();
                    _ended = true;
                    return 0;
                }
                return 2;
            }

            // --- Cell decoding ---
            // All cell records: col (u32 @ 0) + styleAndFlags (u32 @ 4); iStyleRef = low 24 bits.

            private void ProcessCell(int id, ReadOnlySpan<byte> payload)
            {
                if (payload.Length < 8)
                {
                    return;
                }
                int col = (int)Biff12.ReadU32(payload, 0);
                int style = (int)(Biff12.ReadU32(payload, 4) & 0x00FFFFFF);

                switch (id)
                {
                    case Brt.CellRk when payload.Length >= 12:
                        AddDouble(col, style, Biff12.Rk(Biff12.ReadU32(payload, 8)));
                        break;
                    case Brt.CellReal when payload.Length >= 16:
                        AddDouble(col, style, Biff12.ReadF64(payload, 8));
                        break;
                    case Brt.CellIsst when payload.Length >= 12:
                        var (start, len) = _reader.SharedAt((int)Biff12.ReadU32(payload, 8));
                        _acc.Add(col, start, len, CellType.ExcelString, style, CellValueSource.Shared);
                        break;

                    case Brt.CellSt when Biff12.TryReadWideString(payload, 8, out ReadOnlySpan<char> chars, out _):
                        AppendString(col, style, chars);
                        break;
                    case Brt.CellBool when payload.Length >= 9:
                        AppendBool(col, style, payload[8]);
                        break;
                    case Brt.CellError when payload.Length >= 9:
                        AppendError(col, style, payload[8]);
                        break;
                    // Formula cells: the cached result immediately follows the col/style header, in
                    // the same shape as the equivalent plain-cell record; the formula bytes that follow
                    // are outside the record framing we care about (TryReadRecord already bounds payload).
                    case Brt.FmlaNum when payload.Length >= 16:
                        AddDouble(col, style, Biff12.ReadF64(payload, 8));
                        break;
                    case Brt.FmlaString when Biff12.TryReadWideString(payload, 8, out ReadOnlySpan<char> fmlaChars, out _):
                        AppendString(col, style, fmlaChars);
                        break;
                    case Brt.FmlaBool when payload.Length >= 9:
                        AppendBool(col, style, payload[8]);
                        break;
                    case Brt.FmlaError when payload.Length >= 9:
                        AppendError(col, style, payload[8]);
                        break;
                    case Brt.CellRString when payload.Length >= 9 && Biff12.TryReadWideString(payload, 9, out ReadOnlySpan<char> richChars, out _):
                        AppendString(col, style, richChars);
                        break;
                        // CellBlank: no value to emit
                }
            }

            private static bool IsEndSheetData(int id)
            {
                return id == Brt.EndSheetData;
            }

            private void AddDouble(int col, int style, double value)
            {
                CellType type = _reader.IsDateStyle(style) ? CellType.Date : CellType.Number;
                _acc.Add(col, _acc.ValueLength, 0, type, style, CellValueSource.RowValues, number: value, hasNumber: true);
            }

            private void AppendString(int col, int style, ReadOnlySpan<char> chars)
            {
                // Reserve the UTF-8 worst case and encode once, instead of a separate GetByteCount pass.
                int start = _acc.ValueLength;
                Span<byte> dst = _acc.ReserveValueSpan(Encoding.UTF8.GetMaxByteCount(chars.Length));
                _acc.Advance(Encoding.UTF8.GetBytes(chars, dst));
                _acc.Add(col, start, _acc.ValueLength - start, CellType.ExcelString, style, CellValueSource.RowValues);
            }

            private void AppendBool(int col, int style, byte value)
            {
                _acc.AddBool(col, style, value);
            }

            private void AppendError(int col, int style, byte error)
            {
                _acc.AddError(col, style, error);
            }

            private void ResetRow()
            {
                _acc.Reset();
            }

            // --- Binary streaming ---

            private bool TryNextRecord(out int id, out ReadOnlySpan<byte> payload)
            {
                while (true)
                {
                    if (TryNextRecordFromBuffer(out id, out payload))
                    {
                        return true;
                    }
                    if (_eof)
                    {
                        return false;
                    }
                    Fill();
                }
            }

            // Tries to decode one record from the current buffer window. Position advances only on success.
            private bool TryNextRecordFromBuffer(out int id, out ReadOnlySpan<byte> payload)
            {
                var reader = new Biff12RecordReader(_buf.AsSpan(_pos, _len - _pos));
                if (reader.TryReadRecord(out id, out payload))
                {
                    _pos += reader.Position;
                    return true;
                }
                id = -1;
                payload = default;
                return false;
            }

            // Called when the record loop hits EOF unable to decode another record. Unconsumed bytes at
            // that point (_pos < _len) are a record whose framing/payload ran past the end of the stream —
            // i.e. a truncated part. TryReadRecord leaves _pos at that record's start, so the leftover is
            // exactly the partial record. Surface it instead of silently returning the rows read so far.
            private void ThrowIfTruncated()
            {
                if (_pos < _len)
                {
                    throw new InvalidDataException(
                        $"Truncated XLSB worksheet stream: {_len - _pos} trailing byte(s) do not form a complete record.");
                }
            }

            /// <inheritdoc/>
            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
                Justification = "_sheet is opened for this enumerator and owned by it.")]
            public void Dispose()
            {
                _source?.Dispose();
                _source = null;
                ReturnBuffers();
            }

            /// <inheritdoc/>
            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
                Justification = "_sheet is opened for this enumerator and owned by it.")]
            public async ValueTask DisposeAsync()
            {
                if (_source is not null)
                {
                    await _source.DisposeAsync().ConfigureAwait(false);
                    _source = null;
                }
                ReturnBuffers();
            }

        }
    }
}
