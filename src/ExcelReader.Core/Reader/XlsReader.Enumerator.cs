using System.Buffers;
using System.Buffers.Binary;
using System.Buffers.Text;
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
{
    public sealed partial class XlsReader
    {
        [SuppressMessage("Design", "CA1034:Nested types should not be visible",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public sealed class Enumerator : IDisposable, IAsyncDisposable
        {
            private const int InitialVals = 4 * 1024;
            private const int InitialCells = 32;

            [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
                Justification = "Borrowed reader; caller owns its lifetime.")]
            private readonly XlsReader _reader;
            private readonly CancellationToken _ct;
            private int _pos;
            private bool _ended;
            private byte[] _vals;
            private int _valLen;
            private CellDesc[] _cells;
            private int _cellCount;
            private int _row;
            private int _lastCol;
            private bool _sorted;

            internal Enumerator(XlsReader reader, int sheetOffset, CancellationToken ct = default)
            {
                _reader = reader;
                _ct = ct;
                _pos = sheetOffset;
                _row = -1;
                _lastCol = -1;
                _sorted = true;
                _vals = ArrayPool<byte>.Shared.Rent(InitialVals);
                _cells = ArrayPool<CellDesc>.Shared.Rent(InitialCells);
            }

            public Row Current =>
                new(_cells.AsSpan(0, _cellCount), _vals.AsSpan(0, _valLen), _reader.SharedSpan);

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
                ReadOnlySpan<byte> workbook = _reader._workbook;
                while (_pos + 4 <= workbook.Length)
                {
                    int recordStart = _pos;
                    int id = ReadU16Inner(workbook, _pos);
                    int len = ReadU16Inner(workbook, _pos + 2);
                    _pos += 4;
                    if (_pos + len > workbook.Length)
                    {
                        _ended = true;
                        return FinishRow();
                    }
                    ReadOnlySpan<byte> data = workbook.Slice(_pos, len);
                    _pos += len;

                    if (id == Rec.Bof)
                    {
                        if (data.Length < 4 || ReadU16Inner(data, 0) != Biff8Version || ReadU16Inner(data, 2) != SubstreamWorksheet)
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
                        _pos = recordStart;
                        return FinishRow();
                    }

                    ParseCellRecord(id, data);
                }

                _ended = true;
                return FinishRow();
            }

            private bool FinishRow()
            {
                if (!_sorted)
                {
                    Array.Sort(_cells, 0, _cellCount, CellDescColumnComparer.Instance);
                    _sorted = true;
                }
                return _cellCount > 0;
            }

            private void ResetRow()
            {
                _cellCount = 0;
                _valLen = 0;
                _row = -1;
                _lastCol = -1;
                _sorted = true;
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
                        row = ReadU16Inner(data, 0);
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
                            int col = ReadU16Inner(data, 2);
                            int style = ReadU16Inner(data, 4);
                            var (start, len) = _reader.SharedAt(ReadI32Inner(data, 6));
                            AddCell(col, start, len, CellType.ExcelString, style, fromShared: true);
                        }
                        break;
                    case Rec.Number:
                        if (data.Length >= 14)
                        {
                            AddDouble(ReadU16Inner(data, 2), ReadU16Inner(data, 4), BinaryPrimitives.ReadDoubleLittleEndian(data.Slice(6, 8)));
                        }
                        break;
                    case Rec.Rk:
                        if (data.Length >= 10)
                        {
                            AddDouble(ReadU16Inner(data, 2), ReadU16Inner(data, 4), DecodeRk(ReadU32(data, 6)));
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
                    default:
                        // Blank / MulBlank and any other record contribute no value.
                        break;
                }
            }

            private void ParseLabel(ReadOnlySpan<byte> data)
            {
                if (data.Length < 9)
                {
                    return;
                }
                int col = ReadU16Inner(data, 2);
                int style = ReadU16Inner(data, 4);
                int chars = ReadU16Inner(data, 6);
                byte flags = data[8];
                const int start = 9;
                int byteCount = (flags & 1) == 0 ? chars : chars * 2;
                if (start + byteCount > data.Length)
                {
                    return;
                }
                int valueStart = _valLen;
                EnsureValsCapacity(_valLen + (chars * 4));
                _valLen += DecodeStringToUtf8(data.Slice(start, byteCount), chars, flags, _vals.AsSpan(_valLen));
                AddCell(col, valueStart, _valLen - valueStart, CellType.ExcelString, style, fromShared: false);
            }

            private void ParseMulRk(ReadOnlySpan<byte> data)
            {
                if (data.Length < 10)
                {
                    return;
                }
                int col = ReadU16Inner(data, 2);
                int end = data.Length - 2;
                for (int pos = 4; pos + 6 <= end; pos += 6, col++)
                {
                    AddDouble(col, ReadU16Inner(data, pos), DecodeRk(ReadU32(data, pos + 2)));
                }
            }

            private void ParseBoolErr(ReadOnlySpan<byte> data)
            {
                if (data.Length < 8)
                {
                    return;
                }
                int col = ReadU16Inner(data, 2);
                int style = ReadU16Inner(data, 4);
                byte value = data[6];
                bool isError = data[7] != 0;
                int start = _valLen;
                if (isError)
                {
                    AppendError(value);
                    AddCell(col, start, _valLen - start, CellType.Error, style, fromShared: false);
                }
                else
                {
                    EnsureValsCapacity(_valLen + 1);
                    _vals[_valLen++] = value == 0 ? (byte)'0' : (byte)'1';
                    AddCell(col, start, 1, CellType.Boolean, style, fromShared: false);
                }
            }

            private void ParseFormula(ReadOnlySpan<byte> data)
            {
                if (data.Length < 20)
                {
                    return;
                }
                int col = ReadU16Inner(data, 2);
                int style = ReadU16Inner(data, 4);
                ReadOnlySpan<byte> result = data.Slice(6, 8);
                int start = _valLen;
                if (result[6] == 0xFF && result[7] == 0xFF)
                {
                    switch (result[0])
                    {
                        case 1:
                            EnsureValsCapacity(_valLen + 1);
                            _vals[_valLen++] = result[2] == 0 ? (byte)'0' : (byte)'1';
                            AddCell(col, start, 1, CellType.Boolean, style, fromShared: false);
                            break;
                        case 2:
                            AppendError(result[2]);
                            AddCell(col, start, _valLen - start, CellType.Error, style, fromShared: false);
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
                int start = _valLen;
                EnsureValsCapacity(_valLen + 32);
                if (!Utf8Formatter.TryFormat(value, _vals.AsSpan(_valLen), out int written))
                {
                    return;
                }
                _valLen += written;
                CellType type = forced ?? (_reader.IsDateStyle(style) ? CellType.Date : CellType.Number);
                AddCell(col, start, written, type, style, fromShared: false);
            }

            private void AppendError(byte error)
            {
                ReadOnlySpan<byte> text = error switch
                {
                    0x00 => "#NULL!"u8,
                    0x07 => "#DIV/0!"u8,
                    0x0F => "#VALUE!"u8,
                    0x17 => "#REF!"u8,
                    0x1D => "#NAME?"u8,
                    0x24 => "#NUM!"u8,
                    0x2A => "#N/A"u8,
                    _ => "#ERR"u8,
                };
                EnsureValsCapacity(_valLen + text.Length);
                text.CopyTo(_vals.AsSpan(_valLen));
                _valLen += text.Length;
            }

            private void AddCell(int col, int start, int len, CellType type, int style, bool fromShared)
            {
                if (_cellCount == _cells.Length)
                {
                    CellDesc[] bigger = ArrayPool<CellDesc>.Shared.Rent(_cells.Length * 2);
                    Array.Copy(_cells, bigger, _cellCount);
                    ArrayPool<CellDesc>.Shared.Return(_cells);
                    _cells = bigger;
                }
                if (col < _lastCol)
                {
                    _sorted = false;
                }
                _lastCol = col;
                _cells[_cellCount++] = new CellDesc
                {
                    Column = col,
                    Start = start,
                    Length = len,
                    Type = type,
                    Style = style,
                    FromShared = fromShared,
                };
            }

            private void EnsureValsCapacity(int needed)
            {
                if (needed <= _vals.Length)
                {
                    return;
                }
                byte[] bigger = ArrayPool<byte>.Shared.Rent(Math.Max(_vals.Length * 2, needed));
                Array.Copy(_vals, bigger, _valLen);
                ArrayPool<byte>.Shared.Return(_vals);
                _vals = bigger;
            }

            private static double DecodeRk(uint rk)
            {
                double value;
                if ((rk & 0x02) != 0)
                {
                    value = unchecked((int)(rk & 0xFFFFFFFC)) >> 2;
                }
                else
                {
                    ulong raw = (ulong)(rk & 0xFFFFFFFC) << 32;
                    value = BitConverter.Int64BitsToDouble(unchecked((long)raw));
                }
                return (rk & 0x01) != 0 ? value / 100.0 : value;
            }

            private static ushort ReadU16Inner(ReadOnlySpan<byte> src, int offset)
            {
                return BinaryPrimitives.ReadUInt16LittleEndian(src.Slice(offset, 2));
            }

            private static uint ReadU32(ReadOnlySpan<byte> src, int offset)
            {
                return BinaryPrimitives.ReadUInt32LittleEndian(src.Slice(offset, 4));
            }

            private static int ReadI32Inner(ReadOnlySpan<byte> src, int offset)
            {
                return BinaryPrimitives.ReadInt32LittleEndian(src.Slice(offset, 4));
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

        private sealed class CellDescColumnComparer : IComparer<CellDesc>
        {
            internal static readonly CellDescColumnComparer Instance = new();

            public int Compare(CellDesc x, CellDesc y)
            {
                return x.Column.CompareTo(y.Column);
            }
        }
    }
}
