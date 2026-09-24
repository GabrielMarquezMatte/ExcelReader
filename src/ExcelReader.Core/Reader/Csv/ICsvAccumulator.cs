namespace ExcelReader.Core.Reader.Csv
{
    /// <summary>
    /// Folds typed CSV records read by <c>CsvParallel.AggregateAsync</c> into a result. Each partition
    /// of the source gets its own instance.
    /// </summary>
    /// <typeparam name="TSelf">The implementing type.</typeparam>
    /// <typeparam name="TModel">The record type, produced by <see cref="ICsvRecord{TSelf}.TryParse"/> or by an <see cref="Parser.ExcelParser{T}"/>.</typeparam>
    /// <remarks>
    /// <para>
    /// <see cref="Add"/> runs concurrently across instances on worker threads, and must mutate only the
    /// instance it is called on. A partition that started at a misguessed record boundary is read again
    /// into a new instance and the first one is discarded, so an instance may see records that are later
    /// thrown away — including records whose parse or <see cref="Add"/> threw, and records misparsed from
    /// the wrong start that appear nowhere in the source. A side effect outside the instance is not undone.
    /// </para>
    /// <para>
    /// <see cref="Merge"/> runs on a single thread, left to right in source order, folding each later
    /// partition into the accumulator of everything before it. It may receive an instance that saw no records.
    /// </para>
    /// </remarks>
    public interface ICsvAccumulator<TSelf, TModel>
        where TSelf : ICsvAccumulator<TSelf, TModel>
        where TModel : allows ref struct
    {
        /// <summary>Folds one record into this instance.</summary>
        /// <param name="model">The record. Spans it holds are valid only for the duration of the call.</param>
        void Add(TModel model);

        /// <summary>Folds the records of the partition that immediately follows this one into this instance.</summary>
        /// <param name="following">The accumulator of the following partition. It is not used again afterwards.</param>
        void Merge(TSelf following);
    }
}
