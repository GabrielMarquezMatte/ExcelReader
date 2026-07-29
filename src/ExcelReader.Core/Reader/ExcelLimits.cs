using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace ExcelReader.Core.Reader
{
    // Excel's own hard per-workbook caps (ECMA-376 / [MS-XLS]). Both the readers and the writers
    // enforce these, so they live in one place: a reader that accepts a column index a writer would
    // reject (or the reverse) silently breaks round-tripping.
    internal static class ExcelLimits
    {
        // A..XFD
        internal const int MaxColumns = 16_384;
        // 1..1048576
        internal const int MaxRows = 1_048_576;
        // Excel's universal per-cell text limit. Also keeps BIFF8's `cch` field (a u16) from
        // truncating: 32,767 fits in 16 bits.
        internal const int MaxCellTextLength = 32_767;

        internal static void ThrowIfColumnOutOfRange(int columnIndex)
        {
            if ((uint)columnIndex >= MaxColumns)
            {
                ThrowColumnLimit(columnIndex);
            }
        }

        [DoesNotReturn]
        [MethodImpl(MethodImplOptions.NoInlining)]
        internal static void ThrowColumnLimit(int columnIndex)
        {
            throw new ExcelLimitExceededException("Columns", MaxColumns, columnIndex + 1L);
        }

        [DoesNotReturn]
        [MethodImpl(MethodImplOptions.NoInlining)]
        internal static void ThrowRowLimit(long attemptedRowCount)
        {
            throw new ExcelLimitExceededException("Rows", MaxRows, attemptedRowCount);
        }

        internal static void ThrowIfCellTextTooLong(int length, string paramName)
        {
            if (length > MaxCellTextLength)
            {
                ThrowCellTextTooLong(length, paramName);
            }
        }

        [DoesNotReturn]
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void ThrowCellTextTooLong(int length, string paramName)
        {
            throw new ArgumentException(
                $"Cell text exceeds Excel's {MaxCellTextLength}-character limit ({length} chars).", paramName);
        }
    }
}
