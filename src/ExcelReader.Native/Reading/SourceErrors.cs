namespace ExcelReader.Native.Reading
{
    /// <summary>The message a source callback leaves through xl_set_source_error before returning -1.</summary>
    internal static class SourceErrors
    {
        [ThreadStatic]
        private static string? _pending;

        internal static void Set(string? message)
        {
            _pending = string.IsNullOrEmpty(message) ? null : message;
        }

        internal static void Clear()
        {
            _pending = null;
        }

        internal static IOException Failed(string where, long got, long requested)
        {
            string? message = _pending;
            _pending = null;
            return got == -1
                ? new IOException(message ?? $"{where} failed.")
                : new IOException($"{where} returned {got} for a request of {requested} bytes; expected 0 to {requested}, or -1 on failure.");
        }
    }
}
