namespace ExcelReader.Core.Reader.Internal
{
    internal sealed class ExcelTableSheet : IExcelSheet
    {
        private readonly ExcelTable _table;

        internal ExcelTableSheet(ExcelTable table)
        {
            _table = table;
        }

        public int Index
        {
            get
            {
                return _table.Sheet.Index;
            }
        }

        public string Name
        {
            get
            {
                return _table.Sheet.Name;
            }
        }

        public ExcelSheetVisibility Visibility
        {
            get
            {
                return _table.Sheet.Visibility;
            }
        }

        public bool IsDate1904
        {
            get
            {
                return _table.Sheet.IsDate1904;
            }
        }

        public IExcelRowEnumerator GetEnumerator()
        {
            return Bound(_table.Sheet.GetEnumerator());
        }

        public IExcelRowEnumerator GetAsyncEnumerator(CancellationToken ct = default)
        {
            return Bound(_table.Sheet.GetAsyncEnumerator(ct));
        }

        internal static void ThrowIfHeaderless(IExcelSheet sheet)
        {
            if (sheet is ExcelTableSheet view && view._table.HeaderRowCount == 0)
            {
                throw new InvalidOperationException(
                    $"Table '{view._table.Name}' has no header row, so its columns cannot be matched by name.");
            }
        }

        private TableRowEnumerator Bound(IExcelRowEnumerator rows)
        {
            return new TableRowEnumerator(rows, _table.FirstRow, _table.LastRow - _table.TotalsRowCount, _table.FirstColumn, _table.LastColumn);
        }
    }
}
