using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// An Excel table (a range formatted with "Format as Table", stored as a ListObject): its name, its
    /// position on its sheet and its columns.
    /// </summary>
    /// <remarks>Row and column positions are zero-based and include the header and totals rows.</remarks>
    public sealed class ExcelTable
    {
        internal ExcelTable(TablePart part, IExcelSheet sheet)
        {
            Name = part.Name;
            Sheet = sheet;
            Ref = TablePart.FormatRange(part.FirstRow, part.FirstColumn, part.LastRow, part.LastColumn);
            FirstRow = part.FirstRow;
            FirstColumn = part.FirstColumn;
            LastRow = part.LastRow;
            LastColumn = part.LastColumn;
            HeaderRowCount = part.HeaderRowCount;
            TotalsRowCount = part.TotalsRowCount;
            ColumnNames = Array.AsReadOnly(part.Columns);
            StyleName = part.StyleName;
        }

        /// <summary>Gets the table's name, as formulas refer to it (its display name).</summary>
        public string Name { get; }

        /// <summary>Gets the sheet that holds the table.</summary>
        public IExcelSheet Sheet { get; }

        /// <summary>Gets the table's range in A1 notation, header and totals rows included, such as <c>C4:E8</c>.</summary>
        public string Ref { get; }

        /// <summary>Gets the zero-based row of the table's first row: its header row, when it has one.</summary>
        public int FirstRow { get; }

        /// <summary>Gets the zero-based index of the table's first column.</summary>
        public int FirstColumn { get; }

        /// <summary>Gets the zero-based row of the table's last row: its totals row, when it has one.</summary>
        public int LastRow { get; }

        /// <summary>Gets the zero-based index of the table's last column.</summary>
        public int LastColumn { get; }

        /// <summary>Gets the number of header rows: 1, or 0 when the table's header row is turned off.</summary>
        public int HeaderRowCount { get; }

        /// <summary>Gets the number of totals rows: 1 when the table shows a totals row, otherwise 0.</summary>
        public int TotalsRowCount { get; }

        /// <summary>Gets the table's column names, in column order.</summary>
        public IReadOnlyList<string> ColumnNames { get; }

        /// <summary>Gets the name of the table's style, such as <c>TableStyleMedium2</c>, or <see langword="null"/> when the table has none.</summary>
        public string? StyleName { get; }
    }
}
