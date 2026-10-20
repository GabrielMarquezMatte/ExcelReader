using System.Runtime.InteropServices;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Native.Reading
{
    /// <summary>
    /// The <c>xl_cell_view</c> array handed to a row callback. The views borrow the row's own buffers, so the
    /// caller pins <see cref="Row.RowValues"/>, <see cref="Row.Shared"/> and <see cref="Row.RowBuffer"/>
    /// around <see cref="Fill"/> and the callback it feeds.
    /// </summary>
    internal sealed unsafe class CellViewBuffer : IDisposable
    {
        private NativeCellView* _views;
        private int _capacity;

        internal NativeCellView* Fill(ReadOnlySpan<CellDesc> cells, byte* rowValues, byte* shared, byte* rowBuffer)
        {
            if (cells.Length > _capacity)
            {
                _capacity = Math.Max(Math.Max(64, _capacity * 2), cells.Length);
                _views = (NativeCellView*)NativeMemory.Realloc(_views, (nuint)(_capacity * sizeof(NativeCellView)));
            }
            for (int index = 0; index < cells.Length; index++)
            {
                ref readonly CellDesc cell = ref cells[index];
                byte* source = cell.Source switch
                {
                    CellValueSource.Shared => shared,
                    CellValueSource.RowBuffer => rowBuffer,
                    _ => rowValues,
                };
                _views[index] = new NativeCellView
                {
                    Column = cell.Column,
                    Type = (int)cell.Type,
                    ValueLength = cell.Length,
                    HasNumber = cell.HasNumber ? 1 : 0,
                    Value = cell.Length == 0 ? null : source + cell.Start,
                    Number = cell.Number,
                };
            }
            return _views;
        }

        public void Dispose()
        {
            NativeMemory.Free(_views);
            _views = null;
            _capacity = 0;
        }
    }
}
