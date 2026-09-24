using System.Runtime.InteropServices;
using ExcelReader.Core.Reader;
using ExcelReader.Native.Reading;

namespace ExcelReader.Native.Csv
{
    internal sealed unsafe class CsvAggregateState : IDisposable
    {
        private byte* _scratch;
        private int _capacity;

        internal nint Native { get; set; }

        internal NativeRow WriteRow(in Row row)
        {
            int cellCount = 0;
            int valueBytes = 0;
            foreach (RowCell cell in row.Cells)
            {
                cellCount++;
                valueBytes = checked(valueBytes + cell.Value.Value.Length + 1);
            }

            if (cellCount == 0)
            {
                return default;
            }

            int cellsBytes = checked(cellCount * sizeof(NativeRowCell));
            EnsureCapacity(checked(cellsBytes + valueBytes));

            NativeRowCell* cells = (NativeRowCell*)_scratch;
            byte* values = _scratch + cellsBytes;
            int index = 0;
            int offset = 0;
            foreach (RowCell cell in row.Cells)
            {
                ReadOnlySpan<byte> bytes = cell.Value.Value;
                bytes.CopyTo(new Span<byte>(values + offset, bytes.Length));
                values[offset + bytes.Length] = 0;
                cells[index] = new NativeRowCell
                {
                    Column = cell.ColumnIndex,
                    Type = (int)cell.Value.Type,
                    ValueLength = bytes.Length,
                    Value = (IntPtr)(values + offset),
                };
                offset = checked(offset + bytes.Length + 1);
                index++;
            }

            return new NativeRow { CellCount = cellCount, Cells = (IntPtr)_scratch };
        }

        internal void ReleaseScratch()
        {
            if (_scratch is not null)
            {
                NativeMemory.Free(_scratch);
                _scratch = null;
                _capacity = 0;
            }
        }

        public void Dispose()
        {
            ReleaseScratch();
        }

        private void EnsureCapacity(int required)
        {
            if (_capacity >= required)
            {
                return;
            }

            int grown = Math.Max(required, _capacity == 0 ? 1024 : _capacity * 2);
            _scratch = (byte*)NativeMemory.Realloc(_scratch, (nuint)grown);
            _capacity = grown;
        }
    }
}
