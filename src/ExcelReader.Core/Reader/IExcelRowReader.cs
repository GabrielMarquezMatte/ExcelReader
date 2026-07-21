using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
{
    public interface IExcelRowReader<TEnumerator>
        where TEnumerator : IExcelRowEnumerator
    {
        bool IsDate1904 { get; }
        TEnumerator GetEnumerator();
        TEnumerator GetAsyncEnumerator();
        ValueTask<TEnumerator> GetAsyncEnumeratorAsync(CancellationToken ct = default);
    }

    // The non-generic reader is the generic one specialized to the interface enumerator, plus a
    // sheet-navigation surface and dispose. Unifying them lets the typed parser drive a
    // format-agnostic reader (Excel.Open) and lets callers walk every sheet without downcasting to
    // the concrete XlsxReader/XlsbReader/XlsReader type.
    public interface IExcelRowReader : IExcelRowReader<IExcelRowEnumerator>, IDisposable, IAsyncDisposable
    {
        string SheetName { get; }
        int SheetCount { get; }
        bool TryMoveToSheet(ReadOnlySpan<char> name);
        void MoveToSheet(int index);
    }

    public interface IExcelRowEnumerator : IDisposable, IAsyncDisposable
    {
        Row Current { get; }
        bool MoveNext();
        ValueTask<bool> MoveNextAsync();
    }
}