using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace ExcelReader.Core.Reader.Internal
{
    internal static class ExcelLimits
    {
        internal const int MaxColumns = 16_384;
        internal const int MaxRows = 1_048_576;
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
