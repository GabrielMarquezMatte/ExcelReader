namespace ExcelReader.Core.Writer
{
    public interface IWorkbookWriter<out TSheet> : IAsyncDisposable
    {
        ValueTask StartAsync(CancellationToken ct = default);
        TSheet AddSheet(string name);
        ValueTask EndAsync(CancellationToken ct = default);
        ValueTask FlushAsync(CancellationToken ct = default);
    }

    public interface ISheetWriter<TRow> : IAsyncDisposable
    {
        ValueTask StartAsync(CancellationToken ct = default);
        ValueTask<TRow> StartRowAsync(CancellationToken ct = default);
        ValueTask EndAsync(CancellationToken ct = default);
    }

    public interface IRowWriter : IAsyncDisposable
    {
        void Write(string? value);
        void Write(bool value);
        void Write(bool? value);
        void Write(DateTime value);
        void Write(DateTime? value);
        void Write(DateOnly value);
        void Write(DateOnly? value);
        void Write(TimeOnly value);
        void Write(TimeOnly? value);
        void Write<T>(T value) where T : IUtf8SpanFormattable;
        void Write<T>(T? value) where T : struct, IUtf8SpanFormattable;
        void Skip(int count = 1);
    }
}