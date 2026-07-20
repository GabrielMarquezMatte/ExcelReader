using System.Buffers;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
{
    // Per-row cell/value accumulator shared by the XLSX/XLSB/XLS worksheet enumerators. Owns two pooled
    // buffers: `Values` (decoded UTF-8 cell text, appended as cells are parsed) and `Cells` (the
    // CellDesc list). Rent once per enumerator, Reset() per row, Return() on dispose. Column-sortedness
    // is tracked as cells are added so the XLS reader (whose cells can arrive out of order) can sort
    // lazily; the XLSX/XLSB readers, whose cells are always ascending, simply never call SortByColumn.
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

        // Reserves at least `additional` free bytes in the value buffer and returns the free tail to
        // write into; the caller must follow with Advance(bytesWritten).
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
            byte[] bigger = ArrayPool<byte>.Shared.Rent(LimitChecks.NextBufferSize(_maxCellBytes, _limitName, _vals.Length, needed));
            Array.Copy(_vals, bigger, ValueLength);
            ArrayPool<byte>.Shared.Return(_vals);
            _vals = bigger;
        }

        // Appends one raw value byte (used for the bool '0'/'1' text).
        internal void AppendByte(byte b)
        {
            EnsureCapacity(ValueLength + 1);
            _vals[ValueLength++] = b;
        }

        // Appends the display text for a BIFF numeric error code and returns its byte length.
        internal int AppendErrorText(byte code)
        {
            ReadOnlySpan<byte> text = BiffErrorText(code);
            EnsureCapacity(ValueLength + text.Length);
            text.CopyTo(_vals.AsSpan(ValueLength));
            ValueLength += text.Length;
            return text.Length;
        }

        internal void Add(int col, int start, int len, CellType type, int style, bool fromShared, double number = 0, bool hasNumber = false)
        {
            if (Count == _cells.Length)
            {
                int capacity = LimitChecks.NextBufferSize(
                    _maxCellBytes,
                    _limitName,
                    _cells.Length,
                    Count + 1,
                    Unsafe.SizeOf<CellDesc>());
                CellDesc[] bigger = ArrayPool<CellDesc>.Shared.Rent(capacity);
                Array.Copy(_cells, bigger, Count);
                ArrayPool<CellDesc>.Shared.Return(_cells);
                _cells = bigger;
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
                FromShared = fromShared,
                Number = number,
                HasNumber = hasNumber,
            };
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
            int[] keys = ArrayPool<int>.Shared.Rent(Count);
            try
            {
                for (int i = 0; i < Count; i++)
                {
                    keys[i] = _cells[i].Column;
                }
                Array.Sort(keys, _cells, 0, Count);
            }
            finally
            {
                ArrayPool<int>.Shared.Return(keys);
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

        // BIFF error code -> Excel display text (shared by the XLS and XLSB binary readers).
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
