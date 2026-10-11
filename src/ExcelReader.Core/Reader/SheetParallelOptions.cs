namespace ExcelReader.Core.Reader
{
    /// <summary>Options for reading a sheet across several threads with <see cref="SheetParallel"/>.</summary>
    public sealed class SheetParallelOptions
    {
        /// <summary>Gets the default options: one thread per processor, no header.</summary>
        public static SheetParallelOptions Default { get; } = new();

        /// <summary>Gets the maximum number of reading threads. <c>0</c>, the default, means <see cref="Environment.ProcessorCount"/>; <c>1</c> reads sequentially.</summary>
        public int DegreeOfParallelism { get; init; }

        /// <summary>
        /// Gets the 1-based number of the header row; it and every row before it are skipped.
        /// <c>0</c>, the default, means the sheet has no header.
        /// </summary>
        public int HeaderRow { get; init; }
    }
}
