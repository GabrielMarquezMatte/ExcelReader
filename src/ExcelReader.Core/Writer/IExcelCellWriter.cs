namespace ExcelReader.Core.Writer
{
    // Write-direction counterpart to IExcelCellConverter<T>. Bind it to a property with the same
    // [ExcelConverter(typeof(MyConverter))] attribute used for reading; a converter that implements
    // both interfaces round-trips a custom type through WorkbookRecordWriter and ExcelParser<T>.
    //
    // A single instance is created once and shared across every row and thread, so implementations
    // must be stateless / thread-safe. Write exactly one cell per call (the row auto-advances columns).
    /// <summary>
    /// Writes a value of type <typeparamref name="T"/> as a cell of a row, for use with the
    /// <c>[ExcelConverter]</c> attribute on a property that <see cref="WorkbookRecordWriter{TSheet,TRow}"/> writes.
    /// Implement this alongside <c>IExcelCellConverter&lt;T&gt;</c> on the same converter type to round-trip
    /// a custom type through both writing and reading. A single instance is shared across every row and
    /// thread, so implementations must be stateless and thread-safe.
    /// </summary>
    /// <typeparam name="T">The property type this converter writes.</typeparam>
    public interface IExcelCellWriter<in T>
    {
        /// <summary>
        /// Writes <paramref name="value"/> as exactly one cell on <paramref name="row"/>; the row's
        /// column position auto-advances as with any other <see cref="IRowWriter"/> write.
        /// </summary>
        /// <param name="row">The row to write the cell to.</param>
        /// <param name="value">The value to write.</param>
        void Write(IRowWriter row, T value);
    }
}
