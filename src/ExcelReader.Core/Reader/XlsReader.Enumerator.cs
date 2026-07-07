using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;
using static ExcelReader.Core.Reader.Biff12;

namespace ExcelReader.Core.Reader
{
    public sealed partial class XlsReader
    {
        [SuppressMessage("Design", "CA1034:Nested types should not be visible",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public sealed class Enumerator : IExcelRowEnumerator
        {
            [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
                Justification = "Borrowed reader; caller owns its lifetime.")]
            private readonly XlsReader _reader;
            private readonly CancellationToken _ct;
            private readonly BiffCursor _cursor;
            private readonly CellAccumulator _acc;
            private bool _ended;
            private int _row;

            internal Enumerator(XlsReader reader, int sheetOffset, CancellationToken ct = default)
            {
                _reader = reader;
                _ct = ct;
                _cursor = reader.OpenCursor(sheetOffset);
                _row = -1;
                _acc = new CellAccumulator(reader._options.MaxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes));
            }

            public Row Current =>
                new(_acc.CellSpan, _acc.ValueSpan, _reader.SharedSpan);

            public bool MoveNext()
            {
                _ct.ThrowIfCancellationRequested();
                return MoveNextCore();
            }

            public ValueTask<bool> MoveNextAsync()
            {
                return new ValueTask<bool>(MoveNext());
            }

            private bool MoveNextCore()
            {
                if (_ended)
                {
                    return false;
                }

                ResetRow();
                BiffCursor cursor = _cursor;
                while (true)
                {
                    long recordStart = cursor.Position;
                    if (!cursor.TryReadRecord(out int id, out ReadOnlySpan<byte> data))
                    {
                        _ended = true;
                        return FinishRow();
                    }

                    if (id == Rec.Bof)
                    {
                        if (data.Length < 4 || ReadU16(data, 0) != Biff8Version || ReadU16(data, 2) != SubstreamWorksheet)
                        {
                            throw new NotSupportedException("Only BIFF8 worksheet streams are supported.");
                        }
                        continue;
                    }
                    if (id == Rec.Eof)
                    {
                        _ended = true;
                        return FinishRow();
                    }
                    if (!TryGetCellRow(id, data, out int row))
                    {
                        continue;
                    }
                    if (_row < 0)
                    {
                        _row = row;
                    }
                    else if (row != _row)
                    {
                        cursor.Position = recordStart;
                        return FinishRow();
                    }

                    ParseCellRecord(id, data);
                }
            }

            private bool FinishRow()
            {
                _acc.SortByColumn();
                return _acc.Count > 0;
            }

            private void ResetRow()
            {
                _acc.Reset();
                _row = -1;
            }

            private static bool TryGetCellRow(int id, ReadOnlySpan<byte> data, out int row)
            {
                row = -1;
                if (data.Length < 2)
                {
                    return false;
                }
                switch (id)
                {
                    case Rec.Label:
                    case Rec.LabelSst:
                    case Rec.Number:
                    case Rec.Rk:
                    case Rec.MulRk:
                    case Rec.BoolErr:
                    case Rec.Formula:
                    case Rec.Blank:
                    case Rec.MulBlank:
                        row = ReadU16(data, 0);
                        return true;
                    default:
                        return false;
                }
            }

            private void ParseCellRecord(int id, ReadOnlySpan<byte> data)
            {
                switch (id)
                {
                    case Rec.Label:
                        ParseLabel(data);
                        break;
                    case Rec.LabelSst:
                        if (data.Length >= 10)
                        {
                            int col = ReadU16(data, 2);
                            int style = ReadU16(data, 4);
                            var (start, len) = _reader.SharedAt(ReadI32(data, 6));
                            _acc.Add(col, start, len, CellType.ExcelString, style, fromShared: true);
                        }
                        break;
                    case Rec.Number:
                        if (data.Length >= 14)
                        {
                            AddDouble(ReadU16(data, 2), ReadU16(data, 4), BinaryPrimitives.ReadDoubleLittleEndian(data.Slice(6, 8)));
                        }
                        break;
                    case Rec.Rk:
                        if (data.Length >= 10)
                        {
                            AddDouble(ReadU16(data, 2), ReadU16(data, 4), Rk(ReadU32(data, 6)));
                        }
                        break;
                    case Rec.MulRk:
                        ParseMulRk(data);
                        break;
                    case Rec.BoolErr:
                        ParseBoolErr(data);
                        break;
                    case Rec.Formula:
                        ParseFormula(data);
                        break;
                    // Blank / MulBlank and unknown records contribute no value.
                }
            }

            private void ParseLabel(ReadOnlySpan<byte> data)
            {
                if (data.Length < 9)
                {
                    return;
                }
                int col = ReadU16(data, 2);
                int style = ReadU16(data, 4);
                int chars = ReadU16(data, 6);
                byte flags = data[8];
                const int start = 9;
                int byteCount = (flags & 1) == 0 ? chars : chars * 2;
                if (start + byteCount > data.Length)
                {
                    return;
                }
                int valueStart = _acc.ValueLength;
                Span<byte> dst = _acc.ReserveValueSpan(chars * 4);
                _acc.Advance(DecodeStringToUtf8(data.Slice(start, byteCount), chars, flags, dst));
                _acc.Add(col, valueStart, _acc.ValueLength - valueStart, CellType.ExcelString, style, fromShared: false);
            }

            private void ParseMulRk(ReadOnlySpan<byte> data)
            {
                if (data.Length < 10)
                {
                    return;
                }
                int col = ReadU16(data, 2);
                int end = data.Length - 2;
                for (int pos = 4; pos + 6 <= end; pos += 6, col++)
                {
                    AddDouble(col, ReadU16(data, pos), Rk(ReadU32(data, pos + 2)));
                }
            }

            private void ParseBoolErr(ReadOnlySpan<byte> data)
            {
                if (data.Length < 8)
                {
                    return;
                }
                int col = ReadU16(data, 2);
                int style = ReadU16(data, 4);
                byte value = data[6];
                bool isError = data[7] != 0;
                int start = _acc.ValueLength;
                if (isError)
                {
                    int len = _acc.AppendErrorText(value);
                    _acc.Add(col, start, len, CellType.Error, style, fromShared: false);
                }
                else
                {
                    _acc.AppendByte(value == 0 ? (byte)'0' : (byte)'1');
                    _acc.Add(col, start, 1, CellType.Boolean, style, fromShared: false);
                }
            }

            private void ParseFormula(ReadOnlySpan<byte> data)
            {
                if (data.Length < 20)
                {
                    return;
                }
                int col = ReadU16(data, 2);
                int style = ReadU16(data, 4);
                ReadOnlySpan<byte> result = data.Slice(6, 8);
                int start = _acc.ValueLength;
                if (result[6] == 0xFF && result[7] == 0xFF)
                {
                    switch (result[0])
                    {
                        case 1:
                            _acc.AppendByte(result[2] == 0 ? (byte)'0' : (byte)'1');
                            _acc.Add(col, start, 1, CellType.Boolean, style, fromShared: false);
                            break;
                        case 2:
                            int len = _acc.AppendErrorText(result[2]);
                            _acc.Add(col, start, len, CellType.Error, style, fromShared: false);
                            break;
                        default:
                            break;
                    }
                    return;
                }
                AddDouble(col, style, BinaryPrimitives.ReadDoubleLittleEndian(result), CellType.Formula);
            }

            private void AddDouble(int col, int style, double value, CellType? forced = null)
            {
                // Store the raw double only — no eager formatting. Cell formats lazily if a caller
                // asks for text (GetString/Value); numeric consumers read the double directly.
                CellType type = forced ?? (_reader.IsDateStyle(style) ? CellType.Date : CellType.Number);
                _acc.Add(col, _acc.ValueLength, 0, type, style, fromShared: false, number: value, hasNumber: true);
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
                _cursor.Dispose();
                _acc.Return();
            }
        }
    }
}
