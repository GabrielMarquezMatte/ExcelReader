using System.Buffers;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
{
    internal sealed class CellAccumulator
    {
        private const int InitialVals = 4 * 1024;
        private const int InitialCells = 32;

        private readonly int _maxCellBytes;
        private readonly string _limitName;
        private byte[] _vals;
        private CellDesc[] _cells;
        private int _lastCol;
        private bool _sorted;

        internal CellAccumulator(int maxCellBytes, string limitName)
        {
            _maxCellBytes = maxCellBytes;
            _limitName = limitName;
            _vals = ArrayPool<byte>.Shared.Rent(InitialVals);
            _cells = ArrayPool<CellDesc>.Shared.Rent(InitialCells);
            _lastCol = -1;
            _sorted = true;
        }

        internal ReadOnlySpan<CellDesc> CellSpan => _cells.AsSpan(0, Count);
        internal ReadOnlySpan<byte> ValueSpan => _vals.AsSpan(0, ValueLength);

        internal int Count { get; private set; }
        internal int ValueLength { get; private set; }

        internal void Reset()
        {
            Count = 0;
            ValueLength = 0;
            _lastCol = -1;
            _sorted = true;
        }

        internal Span<byte> ReserveValueSpan(int additional)
        {
            EnsureCapacity(ValueLength + additional);
            return _vals.AsSpan(ValueLength);
        }

        internal void Advance(int written)
        {
            ValueLength += written;
        }

        internal void EnsureCapacity(int needed)
        {
            if (needed <= _vals.Length)
            {
                return;
            }
            GrowVals(needed);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private void GrowVals(int needed)
        {
            byte[] bigger = ArrayPool<byte>.Shared.Rent(LimitChecks.NextBufferSize(_maxCellBytes, _limitName, _vals.Length, needed));
            Array.Copy(_vals, bigger, ValueLength);
            ArrayPool<byte>.Shared.Return(_vals);
            _vals = bigger;
        }

        internal void AppendByte(byte b)
        {
            EnsureCapacity(ValueLength + 1);
            _vals[ValueLength++] = b;
        }

        internal int AppendErrorText(byte code)
        {
            ReadOnlySpan<byte> text = BiffErrorText(code);
            EnsureCapacity(ValueLength + text.Length);
            text.CopyTo(_vals.AsSpan(ValueLength));
            ValueLength += text.Length;
            return text.Length;
        }

        internal void AddBool(int col, int style, byte value)
        {
            int start = ValueLength;
            AppendByte(value == 0 ? (byte)'0' : (byte)'1');
            Add(col, start, 1, CellType.Boolean, style, CellValueSource.RowValues);
        }

        internal void AddError(int col, int style, byte code)
        {
            int start = ValueLength;
            int length = AppendErrorText(code);
            Add(col, start, length, CellType.Error, style, CellValueSource.RowValues);
        }

        internal CellDesc[] RawCells => _cells;

        internal void ReserveCells(int capacity)
        {
            if (capacity <= _cells.Length)
            {
                return;
            }
            CellDesc[] bigger = ArrayPool<CellDesc>.Shared.Rent(capacity);
            Array.Copy(_cells, bigger, Count);
            ArrayPool<CellDesc>.Shared.Return(_cells);
            _cells = bigger;
        }

        internal void CommitAscending(int count, int lastCol)
        {
            Count = count;
            _lastCol = lastCol;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void Add(int col, int start, int len, CellType type, int style, CellValueSource source, double number = 0, bool hasNumber = false, int sharedIndex = -1)
        {
            if ((uint)col >= ExcelLimits.MaxColumns)
            {
                ExcelLimits.ThrowColumnLimit(col);
            }
            if (Count == _cells.Length)
            {
                GrowCells();
            }
            if (col < _lastCol)
            {
                _sorted = false;
            }
            _lastCol = col;
            _cells[Count++] = new CellDesc
            {
                Column = col,
                Start = start,
                Length = len,
                Type = type,
                Style = style,
                Source = source,
                Number = number,
                HasNumber = hasNumber,
                SharedIndex = sharedIndex,
            };
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private void GrowCells()
        {
            int capacity = LimitChecks.NextBufferSize(_maxCellBytes, _limitName, _cells.Length, Count + 1,
                                                      Unsafe.SizeOf<CellDesc>());
            CellDesc[] bigger = ArrayPool<CellDesc>.Shared.Rent(capacity);
            Array.Copy(_cells, bigger, Count);
            ArrayPool<CellDesc>.Shared.Return(_cells);
            _cells = bigger;
        }

        internal void SortByColumn()
        {
            if (_sorted)
            {
                return;
            }
            if (Count <= 1)
            {
                _sorted = true;
                return;
            }
            Span<CellDesc> cells = _cells.AsSpan(0, Count);
            for (int i = 1; i < cells.Length; i++)
            {
                CellDesc current = cells[i];
                int j = i - 1;
                while (j >= 0 && cells[j].Column > current.Column)
                {
                    cells[j + 1] = cells[j];
                    j--;
                }
                cells[j + 1] = current;
            }
            _sorted = true;
        }

        internal void Return()
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

        internal static ReadOnlySpan<byte> BiffErrorText(byte code)
        {
            return code switch
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
        }

    }
}
