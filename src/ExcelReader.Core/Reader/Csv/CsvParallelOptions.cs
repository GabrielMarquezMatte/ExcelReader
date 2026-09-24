namespace ExcelReader.Core.Reader.Csv
{
    /// <summary>Options for reading a CSV source across several threads.</summary>
    public sealed class CsvParallelOptions
    {
        /// <summary>Gets the default options: one thread per processor, the default dialect, no header.</summary>
        public static CsvParallelOptions Default { get; } = new();

        /// <summary>Gets the maximum number of reading threads. <c>0</c>, the default, means <see cref="Environment.ProcessorCount"/>; <c>1</c> reads sequentially.</summary>
        public int DegreeOfParallelism { get; init; }

        /// <summary>Gets the CSV dialect options. Defaults to <see cref="CsvReaderOptions.Default"/>.</summary>
        public CsvReaderOptions Reader { get; init; } = CsvReaderOptions.Default;

        /// <summary>
        /// Gets the 1-based number of the header record; it and every record before it are skipped.
        /// <c>0</c>, the default, means the source has no header.
        /// </summary>
        public int HeaderRow { get; init; }
    }
}
