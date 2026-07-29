namespace ExcelReader.Core.Writer.Internal
{
    // Shared WriterState guard checks for the three state-tracking workbook writers (XlsxWorkbookWriter,
    // XlsbWorkbookWriter, XlsWorkbookWriter): each StartAsync/AddSheet/EndAsync call site repeats the
    // same "already disposed" / "wrong state" checks, differing only in the writer's type name and
    // the action being guarded.
    internal static class WriterStateGuard
    {
        internal static void ThrowIfEnded(WriterState state, object writer)
        {
            ObjectDisposedException.ThrowIf(state == WriterState.Ended, writer);
        }

        internal static void RequireCreated(WriterState state, string typeName)
        {
            if (state != WriterState.Created)
            {
                throw new InvalidOperationException($"{typeName} has already been started.");
            }
        }

        internal static void RequireStarted(WriterState state, string typeName, string action)
        {
            if (state != WriterState.Started)
            {
                throw new InvalidOperationException($"{typeName} must be started before {action}.");
            }
        }

        // Guards row-writer re-entrancy at the start of a new row — every sheet writer must reject
        // starting row N+1 while row N's writer is still open.
        internal static void RequireNoActiveRowForStart(bool rowActive, string rowWriterTypeName)
        {
            if (rowActive)
            {
                throw new InvalidOperationException($"The previous {rowWriterTypeName} must be disposed before starting a new row.");
            }
        }

        // Guards ending the sheet while its row writer is still open.
        internal static void RequireNoActiveRowForEnd(bool rowActive, string rowWriterTypeName)
        {
            if (rowActive)
            {
                throw new InvalidOperationException($"The active {rowWriterTypeName} must be disposed before ending the sheet.");
            }
        }

        internal static void ValidateSheetName(string name)
        {
            if (name.Length is 0 or > 31)
            {
                throw new ArgumentException("Sheet names must be 1 to 31 characters.", nameof(name));
            }
            if (name.IndexOfAny([':', '\\', '/', '?', '*', '[', ']']) >= 0)
            {
                throw new ArgumentException("Sheet names cannot contain : \\ / ? * [ or ].", nameof(name));
            }
        }
    }
}
