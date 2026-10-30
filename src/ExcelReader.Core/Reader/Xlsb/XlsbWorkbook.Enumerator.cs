using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Zip;

namespace ExcelReader.Core.Reader.Xlsb
{
    public sealed partial class XlsbWorkbook
    {
        /// <summary>Forward-only enumerator over an <see cref="XlsbWorkbook"/> sheet's rows.</summary>
        /// <remarks>
        /// Streams the underlying binary <c>sheetN.bin</c> entry through a refillable pooled buffer;
        /// <c>Biff12RecordReader</c> framing guarantees that a partial record at the buffer boundary
        /// is detected and retried after the next fill.
        /// </remarks>
        [SuppressMessage("Design", "CA1034:Nested types should not be visible",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public sealed class Enumerator : PooledStreamRowEnumerator, IExcelRowEnumerator, IRowIndexedEnumerator
        {
            [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed",
                Justification = "XlsbWorkbook is borrowed; its lifetime is managed by the caller, not this enumerator.")]
            private readonly XlsbWorkbook _reader;
            private readonly bool[] _styleIsDate;
            private readonly int[] _sharedOffsets;
            private readonly Utf8StringCache? _contentCache;
            private readonly ZipEntryRef _entry;
            private bool _ended;
            private bool _pendingRowHdr;
            private bool _trackRowIndex;
            private int _rowIndex = -1;
            private int _nextRowIndex = -1;

            internal Enumerator(XlsbWorkbook reader, Stream sheet, long entryLength = 0, CancellationToken ct = default)
                : base(sheet, reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes), WorkbookLookups.InitialBufferCapacity(entryLength), ownsSource: true, ct)
            {
                _reader = reader;
                _styleIsDate = reader._styleIsDate;
                _sharedOffsets = reader._sharedOffsets;
                _contentCache = reader._options.InternStrings ? new Utf8StringCache() : null;
                _lease = reader.Lifetime;
            }

            internal Enumerator(XlsbWorkbook reader, ZipEntryRef entry, CancellationToken ct)
                : base(reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes), WorkbookLookups.InitialBufferCapacity(entry.UncompressedSize), ct)
            {
                _reader = reader;
                _styleIsDate = reader._styleIsDate;
                _sharedOffsets = reader._sharedOffsets;
                _contentCache = reader._options.InternStrings ? new Utf8StringCache() : null;
                _entry = entry;
                _lease = reader.Lifetime;
            }

            internal Enumerator(XlsbWorkbook reader, ReadOnlyMemory<byte> chunk, bool rowOpen, CancellationToken ct)
                : base(chunk, reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes), ct)
            {
                _reader = reader;
                _styleIsDate = reader._styleIsDate;
                _sharedOffsets = reader._sharedOffsets;
                _contentCache = reader._options.InternStrings ? new Utf8StringCache() : null;
                _pendingRowHdr = rowOpen;
            }

            private protected override Stream OpenSource()
            {
                return _reader._zip!.OpenEntryStream(_entry, _reader._decompressedBytes.ForRead(), _reader._options);
            }

            private protected override async ValueTask<Stream> OpenSourceAsync()
            {
                return await _reader._zip!.OpenEntryStreamAsync(_entry, _reader._decompressedBytes.ForRead(), _reader._options, _ct).ConfigureAwait(false);
            }

            /// <inheritdoc/>
            public Row Current =>
                new(_acc.CellSpan, _acc.ValueSpan, _reader.SharedSpan, rowBuffer: default, _reader.SharedStringCache, _contentCache);

            internal long Offset => _io.BaseOffset + _pos;

            // The last MoveNext stopped just past the next row's header record, so that row is already open.
            internal bool RowOpen => _pendingRowHdr;

            internal bool Ended => _ended;

            int IRowIndexedEnumerator.RowIndex
            {
                get
                {
                    return _rowIndex;
                }
            }

            void IRowIndexedEnumerator.EnableRowIndex()
            {
                _trackRowIndex = true;
            }

            /// <inheritdoc/>
            public bool MoveNext()
            {
                _ct.ThrowIfCancellationRequested();
                return MoveNextCore() || Finish();
            }

            /// <inheritdoc/>
            public ValueTask<bool> MoveNextAsync()
            {
                _ct.ThrowIfCancellationRequested();
                if (_ended)
                {
                    return FinishAsync();
                }
                ResetRow();
                if (!_pendingRowHdr)
                {
                    int seek = SeekRowHdrFromBuffer();
                    if (seek == 0)
                    {
                        return FinishAsync();
                    }
                    if (seek == 2)
                    {
                        return MoveNextRowAsync(seekDone: false);
                    }
                }
                _pendingRowHdr = false;
                _rowIndex = _nextRowIndex;
                if (CollectCellsFromBuffer() == 2)
                {
                    return MoveNextRowAsync(seekDone: true);
                }
                return new ValueTask<bool>(true);
            }

            private async ValueTask<bool> MoveNextRowAsync(bool seekDone)
            {
                if (!seekDone && !await SeekRowHdrAsync().ConfigureAwait(false))
                {
                    return await FinishAsync().ConfigureAwait(false);
                }
                _pendingRowHdr = false;
                _rowIndex = _nextRowIndex;
                await CollectCellsAsync().ConfigureAwait(false);
                return true;
            }

            private bool MoveNextCore()
            {
                if (_ended)
                {
                    return false;
                }
                ResetRow();
                if (!_pendingRowHdr && !SkipToRowHdr())
                {
                    return false;
                }
                _pendingRowHdr = false;
                _rowIndex = _nextRowIndex;
                CollectCells();
                return true;
            }

            private bool SkipToRowHdr()
            {
                while (true)
                {
                    int result = SeekRowHdrFromBuffer();
                    if (result != 2)
                    {
                        return result == 1;
                    }
                    Fill();
                }
            }

            private void CollectCells()
            {
                while (CollectCellsFromBuffer() == 2)
                {
                    Fill();
                }
            }

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

            private int SeekRowHdrFromBuffer()
            {
                Biff12RecordReader reader = new(_buf.AsSpan(_pos, _len - _pos));
                while (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
                {
                    if (id == Brt.RowHdr)
                    {
                        if (_trackRowIndex)
                        {
                            NoteRowHeader(payload);
                        }
                        _pos += reader.Position;
                        return 1;
                    }
                    if (IsEndSheetData(id))
                    {
                        _pos += reader.Position;
                        _ended = true;
                        return 0;
                    }
                }
                _pos += reader.Position;
                if (_eof)
                {
                    ThrowIfTruncated();
                    _ended = true;
                    return 0;
                }
                return 2;
            }

            private int CollectCellsFromBuffer()
            {
                ReadOnlySpan<byte> data = _buf.AsSpan(_pos, _len - _pos);
                int position = 0;
                while (true)
                {
                    var reader = new Biff12RecordReader(data, DrainCells(data, position));
                    bool read = reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload);
                    position = reader.Position;
                    if (!read)
                    {
                        break;
                    }
                    if (id == Brt.RowHdr)
                    {
                        if (_trackRowIndex)
                        {
                            NoteRowHeader(payload);
                        }
                        _pos += position;
                        _pendingRowHdr = true;
                        return 1;
                    }
                    if (IsEndSheetData(id))
                    {
                        _pos += position;
                        _ended = true;
                        return 0;
                    }
                    ProcessCell(id, payload);
                }
                _pos += position;
                if (_eof)
                {
                    ThrowIfTruncated();
                    _ended = true;
                    return 0;
                }
                return 2;
            }

            // Writes the common number and shared-string cell records straight into the cell array.
            // Stops at the first record it is not sure about and leaves it to ProcessCell.
            private int DrainCells(ReadOnlySpan<byte> data, int position)
            {
                CellAccumulator acc = _acc;
                CellDesc[] cells = acc.RawCells;
                bool[] styleIsDate = _styleIsDate;
                int[] sharedOffsets = _sharedOffsets;
                int count = acc.Count;
                int lastCol = acc.LastColumn;
                int numberStart = acc.ValueLength;
                bool sorted = true;
                ref byte start = ref MemoryMarshal.GetReference(data);
                while (position + 2 <= data.Length && count < cells.Length)
                {
                    int id = Unsafe.Add(ref start, position);
                    int length = Unsafe.Add(ref start, position + 1);
                    if (((id | length) & 0x80) != 0 || length < 12 || position + 2 + length > data.Length)
                    {
                        break;
                    }
                    ref byte payload = ref Unsafe.Add(ref start, position + 2);
                    uint col = Biff12.ReadU32(ref payload, 0);
                    if (col >= ExcelLimits.MaxColumns)
                    {
                        break;
                    }
                    int style = (int)(Biff12.ReadU32(ref payload, 4) & 0x00FFFFFF);
                    if (id == Brt.CellIsst)
                    {
                        (int sharedStart, int sharedLength, int sharedIndex) = WorkbookLookups.SharedAt(sharedOffsets, (int)Biff12.ReadU32(ref payload, 8));
                        cells[count] = new CellDesc
                        {
                            Column = (int)col,
                            Start = sharedStart,
                            Length = sharedLength,
                            Type = CellType.ExcelString,
                            Style = style,
                            Source = CellValueSource.Shared,
                            SharedIndex = sharedIndex,
                        };
                    }
                    else
                    {
                        double value;
                        if (id == Brt.CellRk)
                        {
                            value = Biff12.Rk(Biff12.ReadU32(ref payload, 8));
                        }
                        else if ((id == Brt.CellReal || id == Brt.FmlaNum) && length >= 16)
                        {
                            value = Biff12.ReadF64(ref payload, 8);
                        }
                        else
                        {
                            break;
                        }
                        cells[count] = new CellDesc
                        {
                            Column = (int)col,
                            Start = numberStart,
                            Type = (uint)style < (uint)styleIsDate.Length && styleIsDate[style] ? CellType.Date : CellType.Number,
                            Style = style,
                            Source = CellValueSource.RowValues,
                            Number = value,
                            HasNumber = true,
                            SharedIndex = -1,
                        };
                    }
                    sorted &= (int)col >= lastCol;
                    lastCol = (int)col;
                    count++;
                    position += 2 + length;
                }
                acc.Commit(count, lastCol, sorted);
                return position;
            }

            [MethodImpl(MethodImplOptions.NoInlining)]
            private void NoteRowHeader(ReadOnlySpan<byte> payload)
            {
                if (payload.Length >= 4)
                {
                    _nextRowIndex = Biff12.ReadI32(payload, 0);
                    return;
                }
                _nextRowIndex++;
            }

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
                    case Brt.CellReal or Brt.FmlaNum when payload.Length >= 16:
                        AddDouble(col, style, Biff12.ReadF64(payload, 8));
                        break;
                    case Brt.CellIsst when payload.Length >= 12:
                        var (start, len, sharedIndex) = WorkbookLookups.SharedAt(_sharedOffsets, (int)Biff12.ReadU32(payload, 8));
                        _acc.Add(col, start, len, CellType.ExcelString, style, CellValueSource.Shared, sharedIndex: sharedIndex);
                        break;

                    case Brt.CellSt or Brt.FmlaString:
                        AddInlineString(col, style, payload, 8);
                        break;
                    case Brt.CellBool or Brt.FmlaBool when payload.Length >= 9:
                        AppendBool(col, style, payload[8]);
                        break;
                    case Brt.CellError or Brt.FmlaError when payload.Length >= 9:
                        AppendError(col, style, payload[8]);
                        break;
                    case Brt.CellRString when payload.Length >= 9:
                        AddInlineString(col, style, payload, 9);
                        break;
                }
            }

            [MethodImpl(MethodImplOptions.NoInlining)]
            private void AddInlineString(int col, int style, ReadOnlySpan<byte> payload, int offset)
            {
                if (Biff12.TryReadWideString(payload, offset, out ReadOnlySpan<char> chars, out _))
                {
                    AppendString(col, style, chars);
                }
            }

            private static bool IsEndSheetData(int id)
            {
                return id == Brt.EndSheetData;
            }

            private void AddDouble(int col, int style, double value)
            {
                CellType type = WorkbookLookups.IsDateStyle(_styleIsDate, style) ? CellType.Date : CellType.Number;
                CellAccumulator acc = _acc;
                acc.Add(col, acc.ValueLength, 0, type, style, CellValueSource.RowValues, number: value, hasNumber: true);
            }

            private void AppendString(int col, int style, ReadOnlySpan<char> chars)
            {
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


            private void ThrowIfTruncated()
            {
                if (_pos < _len)
                {
                    throw new InvalidDataException(
                        $"Truncated XLSB worksheet stream: {_len - _pos} trailing byte(s) do not form a complete record.");
                }
            }

        }
    }
}
