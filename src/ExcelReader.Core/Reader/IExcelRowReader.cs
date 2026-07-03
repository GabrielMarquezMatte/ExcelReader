using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
{
    public interface IExcelRowReader<TEnumerator>
        where TEnumerator : IExcelRowEnumerator
    {
        bool IsDate1904 { get; }
        TEnumerator GetEnumerator();
        ValueTask<TEnumerator> GetAsyncEnumeratorAsync(CancellationToken ct = default);
    }

    // The non-generic reader is the generic one specialized to the interface enumerator, plus the
    // sheet-navigation surface of IExcelReader (which also carries IDisposable/IAsyncDisposable).
    // Unifying them lets the typed parser drive a format-agnostic reader (Excel.Open) and lets callers
    // walk every sheet without downcasting to the concrete XlsxReader/XlsbReader/XlsReader type.
    public interface IExcelRowReader : IExcelRowReader<IExcelRowEnumerator>, IExcelReader
    {
        // Both bases declare IsDate1904; re-declare it here to unify the two into one member.
        new bool IsDate1904 { get; }
    }

    public interface IExcelRowEnumerator : IDisposable, IAsyncDisposable
    {
        Row Current { get; }
        bool MoveNext();
        ValueTask<bool> MoveNextAsync();
    }
}