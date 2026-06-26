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

    public interface IExcelRowEnumerator : IDisposable, IAsyncDisposable
    {
        Row Current { get; }
        bool MoveNext();
        ValueTask<bool> MoveNextAsync();
    }
}