using System.Buffers.Binary;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native.Reading
{
    /// <summary>
    /// The row <c>xl_next_row_view</c> hands out: cell structs and NUL-terminated values in native memory
    /// owned by the handle, overwritten by the next row and freed with it.
    /// </summary>
    internal sealed unsafe class RowViewBuffer : IDisposable
    {
        private const int MaxFormattedValueLength = 32;

        private NativeRowCell* _cells;
        private int _cellCapacity;
        private byte* _values;
        private int _valueCapacity;

        internal NativeRow Fill(in Row row)
        {
            int count = 0;
            int used = 0;
            foreach (RowCell cell in row.Cells)
            {
                ReadOnlySpan<byte> text = cell.Value.Value;
                EnsureValues(used + (text.IsEmpty ? MaxFormattedValueLength : text.Length) + 1);
                if (!cell.Value.TryFormat(new Span<byte>(_values + used, _valueCapacity - used), out int written))
                {
                    throw new InvalidOperationException("Cell format buffer too small");
                }
                Add(ref count, cell.ColumnIndex, (int)cell.Value.Type, used, written);
                used += written + 1;
            }
            return Finish(count);
        }

        /// <summary>Takes over a row <c>xl_next_row</c> already serialized but could not hand back.</summary>
        internal NativeRow Fill(ReadOnlySpan<byte> blob)
        {
            int cellCount = BinaryPrimitives.ReadInt32LittleEndian(blob);
            int count = 0;
            int used = 0;
            int offset = sizeof(int);
            for (int index = 0; index < cellCount; index++)
            {
                int column = BinaryPrimitives.ReadInt32LittleEndian(blob[offset..]);
                int type = BinaryPrimitives.ReadInt32LittleEndian(blob[(offset + 4)..]);
                int length = BinaryPrimitives.ReadInt32LittleEndian(blob[(offset + 8)..]);
                offset += RowBlob.CellHeaderSize;
                EnsureValues(used + length + 1);
                blob.Slice(offset, length).CopyTo(new Span<byte>(_values + used, length));
                Add(ref count, column, type, used, length);
                offset += length;
                used += length + 1;
            }
            return Finish(count);
        }

        private void Add(ref int count, int column, int type, int valueOffset, int length)
        {
            _values[valueOffset + length] = 0;
            if (count == _cellCapacity)
            {
                _cellCapacity = Math.Max(64, _cellCapacity * 2);
                _cells = (NativeRowCell*)NativeMemory.Realloc(_cells, (nuint)(_cellCapacity * sizeof(NativeRowCell)));
            }
            // Offsets until the row is complete: growing _values moves it, so pointers are fixed up in Finish.
            _cells[count++] = new NativeRowCell { Column = column, Type = type, ValueLength = length, Value = valueOffset };
        }

        private NativeRow Finish(int count)
        {
            for (int index = 0; index < count; index++)
            {
                _cells[index].Value = (nint)(_values + _cells[index].Value);
            }
            return new NativeRow { CellCount = count, Cells = count == 0 ? IntPtr.Zero : (IntPtr)_cells };
        }

        private void EnsureValues(int required)
        {
            if (required <= _valueCapacity)
            {
                return;
            }
            _valueCapacity = Math.Max(Math.Max(4096, _valueCapacity * 2), required);
            _values = (byte*)NativeMemory.Realloc(_values, (nuint)_valueCapacity);
        }

        public void Dispose()
        {
            NativeMemory.Free(_cells);
            NativeMemory.Free(_values);
            _cells = null;
            _values = null;
            _cellCapacity = 0;
            _valueCapacity = 0;
        }
    }
}
