using System.Buffers.Binary;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Native.Reading
{
    /// <summary>
    /// An <c>xl_cell_view</c> array in native memory. <see cref="Borrow"/> points the views at the row's own
    /// buffers, so the caller pins <see cref="Row.RowValues"/>, <see cref="Row.Shared"/> and
    /// <see cref="Row.RowBuffer"/> around it and the callback it feeds. <see cref="Copy(in Row)"/> copies the
    /// text into memory this buffer owns, so the views outlive the row until the next fill.
    /// </summary>
    internal sealed unsafe class CellViewBuffer : IDisposable
    {
        private NativeCellView* _views;
        private int _capacity;
        private byte* _values;
        private int _valueCapacity;

        internal NativeCellView* Borrow(ReadOnlySpan<CellDesc> cells, byte* rowValues, byte* shared, byte* rowBuffer)
        {
            EnsureViews(cells.Length);
            for (int index = 0; index < cells.Length; index++)
            {
                ref readonly CellDesc cell = ref cells[index];
                byte* source = cell.Source switch
                {
                    CellValueSource.Shared => shared,
                    CellValueSource.RowBuffer => rowBuffer,
                    _ => rowValues,
                };
                _views[index] = View(in cell, cell.Length == 0 ? null : source + cell.Start);
            }
            return _views;
        }

        internal NativeCellView* Copy(in Row row, out int count)
        {
            ReadOnlySpan<CellDesc> cells = row.Descriptors;
            count = cells.Length;
            EnsureViews(cells.Length);
            int total = 0;
            foreach (ref readonly CellDesc cell in cells)
            {
                total = checked(total + cell.Length);
            }
            EnsureValues(total);
            byte* next = _values;
            for (int index = 0; index < cells.Length; index++)
            {
                ref readonly CellDesc cell = ref cells[index];
                byte* value = null;
                if (cell.Length != 0)
                {
                    value = next;
                    row.ValueOf(in cell).CopyTo(new Span<byte>(value, cell.Length));
                    next += cell.Length;
                }
                _views[index] = View(in cell, value);
            }
            return count == 0 ? null : _views;
        }

        /// <summary>Takes over a row <c>xl_rows_next</c> already serialized as text but could not hand back.</summary>
        internal NativeCellView* Copy(ReadOnlySpan<byte> blob, out int count)
        {
            count = BinaryPrimitives.ReadInt32LittleEndian(blob);
            EnsureViews(count);
            EnsureValues(blob.Length);
            byte* next = _values;
            int offset = sizeof(int);
            for (int index = 0; index < count; index++)
            {
                int length = BinaryPrimitives.ReadInt32LittleEndian(blob[(offset + 8)..]);
                byte* value = null;
                if (length != 0)
                {
                    value = next;
                    blob.Slice(offset + RowBlob.CellHeaderSize, length).CopyTo(new Span<byte>(value, length));
                    next += length;
                }
                _views[index] = new NativeCellView
                {
                    Column = BinaryPrimitives.ReadInt32LittleEndian(blob[offset..]),
                    Type = BinaryPrimitives.ReadInt32LittleEndian(blob[(offset + 4)..]),
                    ValueLength = length,
                    Value = value,
                };
                offset += RowBlob.CellHeaderSize + length;
            }
            return count == 0 ? null : _views;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static NativeCellView View(in CellDesc cell, byte* value)
        {
            return new NativeCellView
            {
                Column = cell.Column,
                Type = (int)cell.Type,
                ValueLength = cell.Length,
                HasNumber = cell.HasNumber ? 1 : 0,
                Value = value,
                Number = cell.Number,
            };
        }

        private void EnsureViews(int required)
        {
            if (required > _capacity)
            {
                GrowViews(required);
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private void GrowViews(int required)
        {
            _capacity = Math.Max(Math.Max(64, _capacity * 2), required);
            _views = (NativeCellView*)NativeMemory.Realloc(_views, (nuint)(_capacity * sizeof(NativeCellView)));
        }

        private void EnsureValues(int required)
        {
            if (required > _valueCapacity)
            {
                GrowValues(required);
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private void GrowValues(int required)
        {
            _valueCapacity = Math.Max(Math.Max(4096, _valueCapacity * 2), required);
            NativeMemory.Free(_values);
            _values = (byte*)NativeMemory.Alloc((nuint)_valueCapacity);
        }

        public void Dispose()
        {
            NativeMemory.Free(_views);
            NativeMemory.Free(_values);
            _views = null;
            _values = null;
            _capacity = 0;
            _valueCapacity = 0;
        }
    }
}
