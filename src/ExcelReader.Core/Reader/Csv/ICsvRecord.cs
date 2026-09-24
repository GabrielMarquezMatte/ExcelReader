namespace ExcelReader.Core.Reader.Csv
{
    /// <summary>
    /// A typed CSV record that parses itself from a <see cref="Row"/>, for
    /// <see cref="CsvParallel.AggregateAsync{TAccumulator, TRecord}(ReadOnlyMemory{byte}, CsvParallelOptions?, CancellationToken)"/>.
    /// </summary>
    /// <typeparam name="TSelf">The implementing type. May be a <see langword="ref struct"/> holding spans taken from the row.</typeparam>
    public interface ICsvRecord<TSelf>
        where TSelf : ICsvRecord<TSelf>, allows ref struct
    {
        /// <summary>Parses one record.</summary>
        /// <param name="row">The record. Spans taken from it stay valid only until the accumulator's <c>Add</c> returns.</param>
        /// <param name="record">The parsed record.</param>
        /// <returns><see langword="false"/> to skip the record.</returns>
        static abstract bool TryParse(Row row, out TSelf record);
    }
}
