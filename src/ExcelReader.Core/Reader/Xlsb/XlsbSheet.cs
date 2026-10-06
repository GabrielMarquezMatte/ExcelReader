namespace ExcelReader.Core.Reader.Xlsb
{
    /// <summary>One sheet of an <see cref="XlsbWorkbook"/>. A value that holds no resources; each enumerator it creates reads independently.</summary>
    public readonly record struct XlsbSheet : IExcelSheet, IExcelSheet<XlsbWorkbook.Enumerator>
    {
        private readonly XlsbWorkbook? _workbook;

        internal XlsbSheet(XlsbWorkbook workbook, int index, string name, ExcelSheetVisibility visibility)
        {
            _workbook = workbook;
            Index = index;
            Name = name;
            Visibility = visibility;
            IsDate1904 = workbook.IsDate1904;
        }

        /// <inheritdoc/>
        public int Index { get; }

        /// <inheritdoc/>
        public string Name { get; }

        /// <inheritdoc/>
        public ExcelSheetVisibility Visibility { get; }

        /// <inheritdoc/>
        public bool IsDate1904 { get; }

        private XlsbWorkbook Workbook => _workbook ?? throw new InvalidOperationException("This sheet was not obtained from a workbook.");

        /// <inheritdoc/>
        public XlsbWorkbook.Enumerator GetEnumerator()
        {
            return Workbook.OpenSheet(Index);
        }

        /// <inheritdoc/>
        public XlsbWorkbook.Enumerator GetAsyncEnumerator(CancellationToken ct = default)
        {
            return Workbook.OpenSheetAsync(Index, ct);
        }

        IExcelRowEnumerator IExcelSheet<IExcelRowEnumerator>.GetEnumerator()
        {
            return GetEnumerator();
        }

        IExcelRowEnumerator IExcelSheet<IExcelRowEnumerator>.GetAsyncEnumerator(CancellationToken ct)
        {
            return GetAsyncEnumerator(ct);
        }
    }
}
