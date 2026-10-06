namespace ExcelReader.Core.Reader.Csv
{
    /// <summary>The single, unnamed sheet of a <see cref="CsvReader"/>. A value that holds no resources.</summary>
    public readonly record struct CsvSheet : IExcelSheet, IExcelSheet<CsvReader.Enumerator>
    {
        private readonly CsvReader? _reader;

        internal CsvSheet(CsvReader reader)
        {
            _reader = reader;
        }

        /// <inheritdoc/>
        public int Index => 0;

        /// <summary>Gets the sheet name. Always the empty string, since a CSV source has a single, unnamed sheet.</summary>
        public string Name => "";

        /// <summary>Gets the sheet's visibility. Always <see cref="ExcelSheetVisibility.Visible"/>: delimited text has no tab bar to hide from.</summary>
        public ExcelSheetVisibility Visibility => ExcelSheetVisibility.Visible;

        /// <inheritdoc/>
        public bool IsDate1904 => false;

        internal CsvReader Reader => _reader ?? throw new InvalidOperationException("This sheet was not obtained from a workbook.");

        /// <inheritdoc/>
        /// <exception cref="InvalidOperationException">The source is a non-seekable stream that was already enumerated.</exception>
        public CsvReader.Enumerator GetEnumerator()
        {
            return Reader.OpenSheet();
        }

        /// <inheritdoc/>
        /// <exception cref="InvalidOperationException">The source is a non-seekable stream that was already enumerated.</exception>
        public CsvReader.Enumerator GetAsyncEnumerator(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            return Reader.OpenSheet(ct);
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
