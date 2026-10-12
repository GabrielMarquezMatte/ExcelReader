namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// Folds typed rows read by <c>SheetParallel.AggregateAsync</c> into a result. Each partition
    /// of the sheet gets its own instance.
    /// </summary>
    /// <typeparam name="TSelf">The implementing type.</typeparam>
    /// <typeparam name="TModel">The record type, produced by <see cref="IRowRecord{TSelf}.TryParse"/> or by an <see cref="Parser.ExcelParser{T}"/>.</typeparam>
    /// <remarks>
    /// <para>
    /// <see cref="Add"/> runs concurrently across instances on worker threads, and must mutate only the
    /// instance it is called on. For CSV, a partition that started at a misguessed row boundary is read again
    /// into a new instance and the first one is discarded, so an instance may see rows that are later
    /// thrown away — including rows whose parse or <see cref="Add"/> threw, and rows misparsed from
    /// the wrong start that appear nowhere in the sheet. A side effect outside the instance is not undone.
    /// XLSX and XLSB never re-read a partition.
    /// </para>
    /// <para>
    /// <see cref="Merge"/> runs on a single thread, left to right in sheet order, folding each later
    /// partition into the accumulator of everything before it. It may receive an instance that saw no rows.
    /// </para>
    /// </remarks>
    public interface IRowAccumulator<TSelf, TModel>
        where TSelf : IRowAccumulator<TSelf, TModel>
        where TModel : allows ref struct
    {
        /// <summary>Folds one row into this instance.</summary>
        /// <param name="model">The row. Spans it holds are valid only for the duration of the call.</param>
        void Add(TModel model);

        /// <summary>Folds the rows of the partition that immediately follows this one into this instance.</summary>
        /// <param name="following">The accumulator of the following partition. It is not used again afterwards.</param>
        void Merge(TSelf following);
    }
}
