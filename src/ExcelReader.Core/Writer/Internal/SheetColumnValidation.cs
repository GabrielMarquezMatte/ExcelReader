namespace ExcelReader.Core.Writer.Internal
{
    // Shared SetColumnStyle/SetColumnWidth validation for the three sheet writers (XlsxSheetWriter,
    // XlsbSheetWriter, XlsSheetWriter): both methods were byte-for-byte identical across all three —
    // validate, gate on WriterState, then stash into the caller's own dictionary — differing only in
    // which writer's _state/_owner.StyleCount is consulted and the "must be called before ___" wording
    // (the XLS writer has no async twin worth naming, so it points at Start instead of StartAsync).
    internal static class SheetColumnValidation
    {
        internal static void SetColumnStyle(
            ref Dictionary<int, int>? columnStyles, int columnIndex, int styleId, int styleCount,
            WriterState state, object writer, string beforeMethodName)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfNegative(styleId);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(styleId, styleCount);
            RequireNotStarted(state, writer, beforeMethodName);
            columnStyles ??= [];
            columnStyles[columnIndex] = styleId;
        }

        internal static void SetColumnWidth(
            ref Dictionary<int, double>? columnWidths, int columnIndex, double width,
            WriterState state, object writer, string beforeMethodName)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentOutOfRangeException.ThrowIfNegative(width);
            RequireNotStarted(state, writer, beforeMethodName);
            columnWidths ??= [];
            columnWidths[columnIndex] = width;
        }

        private static void RequireNotStarted(WriterState state, object writer, string beforeMethodName)
        {
            WriterStateGuard.ThrowIfEnded(state, writer);
            if (state != WriterState.Created)
            {
                throw new InvalidOperationException(
                    $"{nameof(SetColumnStyle)}/{nameof(SetColumnWidth)} must be called before {beforeMethodName}.");
            }
        }
    }
}
