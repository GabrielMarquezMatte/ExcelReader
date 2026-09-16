namespace ExcelReader.Core.Writer.Internal
{
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

        internal static void RequireNoActiveRowForStart(bool rowActive, string rowWriterTypeName)
        {
            if (rowActive)
            {
                throw new InvalidOperationException($"The previous {rowWriterTypeName} must be disposed before starting a new row.");
            }
        }

        internal static void RequireNoActiveRowForEnd(bool rowActive, string rowWriterTypeName)
        {
            if (rowActive)
            {
                throw new InvalidOperationException($"The active {rowWriterTypeName} must be disposed before ending the sheet.");
            }
        }

        internal static void RequireCanAddSheet(
            WriterState state, object owner, string workbookTypeName, string name,
            bool sheetActive, string sheetWriterTypeName)
        {
            ArgumentNullException.ThrowIfNull(name);
            ThrowIfEnded(state, owner);
            RequireStarted(state, workbookTypeName, "adding sheets");
            ValidateSheetName(name);
            if (sheetActive)
            {
                throw new InvalidOperationException(
                    $"The previous {sheetWriterTypeName} must be ended before adding a new sheet.");
            }
        }

        private static void ValidateSheetName(string name)
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
