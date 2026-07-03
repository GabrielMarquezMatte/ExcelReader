namespace ExcelReader.Core.Writer
{
    // Write-direction counterpart to IExcelCellConverter<T>. Bind it to a property with the same
    // [ExcelConverter(typeof(MyConverter))] attribute used for reading; a converter that implements
    // both interfaces round-trips a custom type through WorkbookRecordWriter and ExcelParser<T>.
    //
    // A single instance is created once and shared across every row and thread, so implementations
    // must be stateless / thread-safe. Write exactly one cell per call (the row auto-advances columns).
    public interface IExcelCellWriter<in T>
    {
        void Write(IRowWriter row, T value);
    }
}
