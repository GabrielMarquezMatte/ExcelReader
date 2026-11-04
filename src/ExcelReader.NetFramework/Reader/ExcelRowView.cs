using System.Runtime.InteropServices;
using ExcelReader.NetFramework.Interop;

namespace ExcelReader.NetFramework.Reader
{
    /// <summary>
    /// A row borrowed from the native reader, valid only while the callback that received it runs. Use
    /// <see cref="ToRow"/> to keep a copy.
    /// </summary>
    [StructLayout(LayoutKind.Auto)]
    public readonly unsafe ref struct ExcelRowView
    {
        private readonly XlCellView* _cells;

        internal ExcelRowView(XlCellView* cells, int count)
        {
            _cells = cells;
            Count = count;
        }

        /// <summary>The number of non-empty cells the native reader reported.</summary>
        public int Count { get; }

        /// <summary>The cell at <paramref name="index"/>, in column order.</summary>
        /// <param name="index">The zero-based position of the cell in the row, not its column.</param>
        /// <returns>The cell.</returns>
        public ExcelCellView this[int index]
        {
            get
            {
                if ((uint)index >= (uint)Count)
                {
                    throw new ArgumentOutOfRangeException(nameof(index), index, "The row has fewer cells.");
                }

                return new ExcelCellView(_cells + index);
            }
        }

        /// <summary>Copies the row into an <see cref="ExcelRow"/> that outlives the callback.</summary>
        /// <returns>The copy.</returns>
        public ExcelRow ToRow()
        {
            ExcelCell[] cells = new ExcelCell[Count];
            for (int i = 0; i < cells.Length; i++)
            {
                ExcelCellView cell = new ExcelCellView(_cells + i);
                cells[i] = new ExcelCell(cell.Column, cell.Type, cell.GetText());
            }

            return new ExcelRow(cells);
        }
    }
}
