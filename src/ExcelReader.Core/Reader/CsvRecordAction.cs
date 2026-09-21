namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// Receives one CSV record parsed by <c>Excel.ForEachCsvParallelAsync</c>.
    /// </summary>
    /// <typeparam name="TModel">The record type. May be a <see langword="ref struct"/> holding spans taken from the record.</typeparam>
    /// <param name="model">The record. Spans it holds are valid only for the duration of the call.</param>
    /// <remarks>
    /// Runs concurrently on worker threads and must be safe to call from several at once; records arrive
    /// in no particular order. May be invoked more than once for the same record — see
    /// <see cref="Excel.ForEachCsvParallelAsync{TRecord}(string, CsvRecordAction{TRecord}, CsvParallelOptions?, CancellationToken)"/>.
    /// </remarks>
    public delegate void CsvRecordAction<TModel>(TModel model)
        where TModel : allows ref struct;
}
