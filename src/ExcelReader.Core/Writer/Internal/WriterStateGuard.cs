using ExcelReader.Core.Enums;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class WriterStateGuard
    {
        internal static void ThrowIfEnded(WriterState state, object writer)
        {
            ObjectDisposedException.ThrowIf(state == WriterState.Ended, writer);
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
            bool ended, object owner, string name,
            bool sheetActive, string sheetWriterTypeName, ExcelSheetVisibility visibility = ExcelSheetVisibility.Visible)
        {
            ArgumentNullException.ThrowIfNull(name);
            ObjectDisposedException.ThrowIf(ended, owner);
            ValidateSheetName(name);
            if (visibility is not (ExcelSheetVisibility.Visible or ExcelSheetVisibility.Hidden or ExcelSheetVisibility.VeryHidden))
            {
                throw new ArgumentOutOfRangeException(nameof(visibility), visibility, "Not a defined ExcelSheetVisibility value.");
            }
            if (sheetActive)
            {
                throw new InvalidOperationException(
                    $"The previous {sheetWriterTypeName} must be ended before adding a new sheet.");
            }
        }

        /// <summary>
        /// Guards against a workbook whose sheets are all hidden: the file formats allow it, but Excel
        /// reports such a workbook as damaged, so producing one is a bug worth surfacing at write time.
        /// </summary>
        internal static void RequireVisibleSheet(bool anyVisible, string workbookTypeName)
        {
            if (!anyVisible)
            {
                throw new InvalidOperationException(
                    $"{workbookTypeName} must contain at least one visible sheet; Excel rejects a workbook whose sheets are all hidden.");
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
