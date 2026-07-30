using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class CellValueGuards
    {
        // NaN/Infinity have no representation in any of the four output formats' numeric cell encodings
        // (XLSX's numeric <v> per ISO/IEC 29500 ST_Xstring, XLSB's Xnum per [MS-XLSB] 2.5.166.6, BIFF8's
        // Xnum per [MS-XLS] 2.5.240). Writing the raw bit pattern or the literal text "NaN" produces a
        // file a conformant reader must reject, so every writer refuses the value here instead.
        internal static void ThrowIfNonFinite(double value, string paramName)
        {
            if (!double.IsFinite(value))
            {
                ThrowNonFinite(value, paramName);
            }
        }

        internal static void ThrowIfNonFinite(float value, string paramName)
        {
            if (!float.IsFinite(value))
            {
                ThrowNonFinite(value, paramName);
            }
        }

        [DoesNotReturn]
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void ThrowNonFinite(double value, string paramName)
        {
            throw new ArgumentException($"Cannot write non-finite value '{value}' to a spreadsheet cell.", paramName);
        }

        [DoesNotReturn]
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void ThrowNonFinite(float value, string paramName)
        {
            throw new ArgumentException($"Cannot write non-finite value '{value}' to a spreadsheet cell.", paramName);
        }
    }
}
