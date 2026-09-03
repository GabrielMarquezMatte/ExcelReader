using System.Diagnostics.CodeAnalysis;
using System.Globalization;
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

        // Every fixed-width numeric type in the BCL formats to a finite decimal literal and converts to a
        // finite double, so the two checks below are dead weight for them. This folds to a JIT constant
        // per instantiation, eliding those checks entirely; only a caller-defined IUtf8SpanFormattable
        // pays for them.
        internal static bool IsAlwaysFinite<T>()
        {
            return typeof(T) == typeof(int) || typeof(T) == typeof(long) || typeof(T) == typeof(short)
                || typeof(T) == typeof(byte) || typeof(T) == typeof(uint) || typeof(T) == typeof(ulong)
                || typeof(T) == typeof(ushort) || typeof(T) == typeof(sbyte) || typeof(T) == typeof(nint)
                || typeof(T) == typeof(nuint) || typeof(T) == typeof(decimal)
                || typeof(T) == typeof(Int128) || typeof(T) == typeof(UInt128);
        }

        // The generic Write<T> path formats a caller's type straight into the cell without ever holding a
        // double, so a type whose TryFormat emits "1e400", "Infinity", or plain text would otherwise put a
        // value into <v> that no conformant reader can represent as a number.
        internal static void ThrowIfNotFiniteNumberText(ReadOnlySpan<byte> utf8, Type sourceType, string paramName)
        {
            if (!double.TryParse(utf8, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                || !double.IsFinite(parsed))
            {
                ThrowUnrepresentable(sourceType, paramName);
            }
        }

        // Reached when a conversion to double overflows or produces NaN. Names the source type: the bare
        // "non-finite value 8" the double overload reports gives no hint which value caused it.
        internal static void ThrowIfNonFiniteConversion(double value, Type sourceType, string paramName)
        {
            if (!double.IsFinite(value))
            {
                ThrowUnrepresentable(sourceType, paramName);
            }
        }

        [DoesNotReturn]
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void ThrowUnrepresentable(Type sourceType, string paramName)
        {
            throw new ArgumentException(
                $"Cannot write a value of type '{sourceType}' to a spreadsheet cell: it does not convert to a finite number.",
                paramName);
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
