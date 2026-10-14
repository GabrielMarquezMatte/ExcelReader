namespace ExcelReader.Core.Writer.Parallel
{
    /// <summary>A record source that is either synchronous or asynchronous.</summary>
    internal readonly struct RecordSource<T>
    {
        internal RecordSource(IEnumerable<T> records)
        {
            Sync = records;
            Async = null;
        }

        internal RecordSource(IAsyncEnumerable<T> records)
        {
            Sync = null;
            Async = records;
        }

        internal IEnumerable<T>? Sync { get; }

        internal IAsyncEnumerable<T>? Async { get; }
    }
}
