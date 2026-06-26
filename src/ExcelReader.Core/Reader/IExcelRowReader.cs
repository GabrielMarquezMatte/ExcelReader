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

    // The non-generic reader is the generic one specialized to the interface enumerator, plus
    // disposal. Unifying them lets the typed parser drive a format-agnostic reader (Excel.Open).
    public interface IExcelRowReader : IExcelRowReader<IExcelRowEnumerator>, IDisposable, IAsyncDisposable
    {
    }

    public interface IExcelRowEnumerator : IDisposable, IAsyncDisposable
    {
        Row Current { get; }
        bool MoveNext();
        ValueTask<bool> MoveNextAsync();
    }
}