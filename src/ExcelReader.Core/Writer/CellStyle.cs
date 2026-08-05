namespace ExcelReader.Core.Writer
{
    /// <summary>A column- or row-level cell style: number format, bold, and italic. Register one with <see cref="IWorkbookWriter{TSheet}.AddStyle"/>.</summary>
    public readonly record struct CellStyle
    {
        /// <summary>Gets the Excel number-format code (e.g. <c>"#,##0.00"</c>, <c>"yyyy-mm-dd"</c>, <c>"0.0%"</c>), or <see langword="null"/> for the general format.</summary>
        public string? NumberFormat { get; init; }

        /// <summary>Gets a value indicating whether text in cells using this style is bold.</summary>
        public bool Bold { get; init; }

        /// <summary>Gets a value indicating whether text in cells using this style is italic.</summary>
        public bool Italic { get; init; }
    }
}
