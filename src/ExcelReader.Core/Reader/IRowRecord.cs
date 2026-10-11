namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// A typed row that parses itself from a <see cref="Row"/>, for
    /// <see cref="SheetParallel.AggregateAsync{TAccumulator, TRecord}(IExcelSheet, SheetParallelOptions?, CancellationToken)"/>.
    /// </summary>
    /// <typeparam name="TSelf">The implementing type. May be a <see langword="ref struct"/> holding spans taken from the row.</typeparam>
    public interface IRowRecord<TSelf>
        where TSelf : IRowRecord<TSelf>, allows ref struct
    {
        /// <summary>Parses one row.</summary>
        /// <param name="row">The row. Spans taken from it stay valid only until the accumulator's <c>Add</c> returns.</param>
        /// <param name="record">The parsed row.</param>
        /// <returns><see langword="false"/> to skip the row.</returns>
        static abstract bool TryParse(Row row, out TSelf record);
    }
}
