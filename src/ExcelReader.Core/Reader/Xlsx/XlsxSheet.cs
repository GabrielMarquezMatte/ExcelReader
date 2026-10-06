namespace ExcelReader.Core.Reader.Xlsx
{
    /// <summary>One sheet of an <see cref="XlsxWorkbook"/>. A value that holds no resources; each enumerator it creates reads independently.</summary>
    public readonly record struct XlsxSheet : IExcelSheet, IExcelSheet<XlsxWorkbook.Enumerator>
    {
        private readonly XlsxWorkbook? _workbook;

        internal XlsxSheet(XlsxWorkbook workbook, int index, string name, ExcelSheetVisibility visibility)
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

        private XlsxWorkbook Workbook => _workbook ?? throw new InvalidOperationException("This sheet was not obtained from a workbook.");

        /// <inheritdoc/>
        public XlsxWorkbook.Enumerator GetEnumerator()
        {
            return Workbook.OpenSheet(Index);
        }

        /// <inheritdoc/>
        public XlsxWorkbook.Enumerator GetAsyncEnumerator(CancellationToken ct = default)
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
