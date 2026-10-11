namespace ExcelReader.Core.Reader
{
    /// <summary>Options for adapting a sheet to <see cref="System.Data.IDataReader"/> with <see cref="ExcelDataReader"/>.</summary>
    public sealed class ExcelDataReaderOptions
    {
        /// <summary>Gets the default options: column names on row 1.</summary>
        public static ExcelDataReaderOptions Default { get; } = new();

        /// <summary>
        /// Gets the 1-based row holding column names. <c>0</c> means a header-less sheet, whose columns come
        /// back named <c>"Column0"</c>, <c>"Column1"</c>, ... and whose count is taken from the first data row.
        /// Defaults to <c>1</c>.
        /// </summary>
        public int HeaderRow { get; init; } = 1;
    }
}
