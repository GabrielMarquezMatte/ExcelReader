namespace ExcelReader.NetFramework.Reader
{
    /// <summary>The name and range of an Excel table. Rows and columns are zero-based and include the header and totals rows.</summary>
    public sealed class ExcelTableInfo
    {
        internal ExcelTableInfo(int index, string name, int sheetIndex, int firstRow, int lastRow, int firstColumn, int lastColumn, int headerRowCount, int totalsRowCount)
        {
            Index = index;
            Name = name;
            SheetIndex = sheetIndex;
            FirstRow = firstRow;
            LastRow = lastRow;
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
            HeaderRowCount = headerRowCount;
            TotalsRowCount = totalsRowCount;
        }

        /// <summary>The zero-based position of the table in the workbook.</summary>
        public int Index { get; }

        /// <summary>The table name.</summary>
        public string Name { get; }

        /// <summary>The zero-based index of the sheet holding the table.</summary>
        public int SheetIndex { get; }

        /// <summary>The first row of the table.</summary>
        public int FirstRow { get; }

        /// <summary>The last row of the table.</summary>
        public int LastRow { get; }

        /// <summary>The first column of the table.</summary>
        public int FirstColumn { get; }

        /// <summary>The last column of the table.</summary>
        public int LastColumn { get; }

        /// <summary>The number of header rows.</summary>
        public int HeaderRowCount { get; }

        /// <summary>The number of totals rows.</summary>
        public int TotalsRowCount { get; }
    }
}
