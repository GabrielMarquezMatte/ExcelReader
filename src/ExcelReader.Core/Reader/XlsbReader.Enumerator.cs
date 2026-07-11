using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
{
    public sealed partial class XlsbReader
    {
        // Forward-only worksheet scanner over a binary sheetN.bin entry. Streams the part through a
        // refillable pooled buffer; Biff12RecordReader framing guarantees that a partial record at the
        // buffer boundary is detected and retried after the next fill.
        [SuppressMessage("Design", "CA1034:Nested types should not be visible",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public sealed class Enumerator : IExcelRowEnumerator
        {
            [SuppressMessage("SharpSource", "SS066:Disposable field is not disposed", Justification = "Borrowed, not owned.")]
            private readonly XlsbReader _reader;
            private readonly CancellationToken _ct;
            [SuppressMessage("SharpSource", "SS066:Disposable field is not disposed", Justification = "Disposed in Dispose().")]
            private Stream? _sheet;
            private readonly BufferedStreamCursor _io;
            private byte[] _buf => _io.Buf;
            private int _pos { get => _io.Pos; set => _io.Pos = value; }
            private int _len => _io.Len;
            private bool _eof => _io.Eof;
            private bool _ended;
            // A BrtRowHdr for the NEXT row was already consumed while collecting cells for the current row.
            // On the next MoveNext call, skip the "seek to row header" step.
            private bool _pendingRowHdr;

            private readonly CellAccumulator _acc;

            internal Enumerator(XlsbReader reader, Stream sheet, CancellationToken ct = default)
            {
                _reader = reader;
                _sheet = sheet;
                _ct = ct;
                _io = new BufferedStreamCursor(reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes));
                _acc = new CellAccumulator(reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes));
            }

            public Row Current =>
                new(_acc.CellSpan, _acc.ValueSpan, _reader.SharedSpan);

            public bool MoveNext()
            {
                _ct.ThrowIfCancellationRequested();
                return MoveNextCore();
            }

            // Non-async fast path mirroring MoveNextCore: with a 64KB buffer, almost every row's
            // header and cells are already fully buffered, so this runs the same *FromBuffer primitives
            // MoveNextCore uses, entirely synchronously, and only pays for an async state machine when
            // a primitive actually reports a buffer miss (result == 2) — a real refill, not per row.
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
                        break;
                    }
                }
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
                    {
                        var (start, len) = _reader.SharedAt((int)Biff12.ReadU32(payload, 8));
                        _acc.Add(col, start, len, CellType.ExcelString, style, fromShared: true);
                        break;
                    }
                    case Brt.CellSt when Biff12.TryReadWideString(payload, 8, out ReadOnlySpan<char> chars, out _):
                        AppendString(col, style, chars);
                        break;
                    case Brt.CellBool when payload.Length >= 9:
                        AppendBool(col, style, payload[8]);
                        break;
                    case Brt.CellError when payload.Length >= 9:
                        AppendError(col, style, payload[8]);
                        break;
                    // CellBlank: no value to emit
                }
            }

            private static bool IsEndSheetData(int id)
            {
                return id is Brt.EndSheetData or Brt.LegacyEndSheetData;
            }

            private void AddDouble(int col, int style, double value)
            {
                CellType type = _reader.IsDateStyle(style) ? CellType.Date : CellType.Number;
                _acc.Add(col, _acc.ValueLength, 0, type, style, fromShared: false, number: value, hasNumber: true);
            }

            private void AppendString(int col, int style, ReadOnlySpan<char> chars)
            {
                // Reserve the UTF-8 worst case and encode once, instead of a separate GetByteCount pass.
                int start = _acc.ValueLength;
                Span<byte> dst = _acc.ReserveValueSpan(Encoding.UTF8.GetMaxByteCount(chars.Length));
                _acc.Advance(Encoding.UTF8.GetBytes(chars, dst));
                _acc.Add(col, start, _acc.ValueLength - start, CellType.ExcelString, style, fromShared: false);
            }

            private void AppendBool(int col, int style, byte value)
            {
                int start = _acc.ValueLength;
                _acc.AppendByte(value == 0 ? (byte)'0' : (byte)'1');
                _acc.Add(col, start, 1, CellType.Boolean, style, fromShared: false);
            }

            private void AppendError(int col, int style, byte error)
            {
                int start = _acc.ValueLength;
                int len = _acc.AppendErrorText(error);
                _acc.Add(col, start, len, CellType.Error, style, fromShared: false);
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

            private void Fill()
            {
                _io.Fill(_sheet!);
            }

            private ValueTask FillAsync()
            {
                return _io.FillAsync(_sheet!, _ct);
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
                _io.Return();
                _acc.Return();
            }
        }
    }
}
