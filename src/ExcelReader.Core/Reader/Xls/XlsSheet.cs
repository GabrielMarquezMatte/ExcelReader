namespace ExcelReader.Core.Reader.Xls
{
    /// <summary>One sheet of an <see cref="XlsWorkbook"/>. A value that holds no resources; each enumerator it creates reads independently.</summary>
    public readonly record struct XlsSheet : IExcelSheet, IExcelSheet<XlsWorkbook.Enumerator>
    {
        private readonly XlsWorkbook? _workbook;

        internal XlsSheet(XlsWorkbook workbook, int index, string name, ExcelSheetVisibility visibility)
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

        private XlsWorkbook Workbook => _workbook ?? throw new InvalidOperationException("This sheet was not obtained from a workbook.");

        /// <inheritdoc/>
        public XlsWorkbook.Enumerator GetEnumerator()
        {
            return Workbook.OpenSheet(Index);
        }

        /// <inheritdoc/>
        public XlsWorkbook.Enumerator GetAsyncEnumerator(CancellationToken ct = default)
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
