using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Writer
{
    /// <summary>Configures the parallel record writes of <see cref="ParallelSheetWriterExtensions"/>.</summary>
    public sealed class ParallelWriteOptions
    {
        /// <summary>Gets the default options: one batch in flight per processor, 8,192 rows per batch.</summary>
        public static ParallelWriteOptions Default { get; } = new();

        /// <summary>
        /// Gets the most batches formatted and compressed at once. <c>0</c>, the default, means
        /// <see cref="Environment.ProcessorCount"/>. With <c>1</c> the batches run on the calling flow,
        /// without the thread pool.
        /// </summary>
        public int DegreeOfParallelism { get; init; }

        /// <summary>
        /// Gets the number of rows per batch. Each batch in flight holds its formatted rows and their
        /// compressed form, so this and <see cref="DegreeOfParallelism"/> bound the memory used.
        /// </summary>
        public int BatchRows { get; init; } = 8192;

        [SuppressMessage("Design", "MA0015:Specify the parameter name in ArgumentException", Justification = "The invalid values are properties of this options object, so the property name is the right parameter name.")]
        internal void Validate()
        {
            ArgumentOutOfRangeException.ThrowIfNegative(DegreeOfParallelism, nameof(DegreeOfParallelism));
            ArgumentOutOfRangeException.ThrowIfLessThan(BatchRows, 1, nameof(BatchRows));
        }
    }
}
