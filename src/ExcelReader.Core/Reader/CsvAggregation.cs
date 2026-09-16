namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// A fold over CSV records expressed as three functions, for
    /// <see cref="Excel.AggregateCsvParallelAsync{TState}(ReadOnlyMemory{byte}, CsvAggregation{TState}, CsvParallelOptions?, CancellationToken)"/>.
    /// Carries the same contract as <see cref="ICsvAccumulator{TSelf, TModel}"/>, without requiring a type.
    /// </summary>
    /// <typeparam name="TState">The accumulator type. Each partition of the source owns one instance.</typeparam>
    public sealed class CsvAggregation<TState>
    {
        /// <summary>Gets the function that creates an empty accumulator. Called once per partition, possibly more.</summary>
        public required Func<TState> Seed { get; init; }

        /// <summary>Gets the function that folds one record into its partition's accumulator. Runs concurrently on worker threads and must mutate only the accumulator it is given.</summary>
        public required CsvRowAction<TState> Accumulate { get; init; }

        /// <summary>Gets the function that folds the following partition's accumulator into the preceding one. Called on one thread, left to right in source order.</summary>
        public required Func<TState, TState, TState> Combine { get; init; }
    }
}
