namespace ExcelReader.Core.Reader
{
    public interface IExcelReader : IDisposable, IAsyncDisposable
    {
        string SheetName { get; }
        int SheetCount { get; }
        bool IsDate1904 { get; }
        bool TryMoveToSheet(ReadOnlySpan<char> name);
        void MoveToSheet(int index);
    }
}
