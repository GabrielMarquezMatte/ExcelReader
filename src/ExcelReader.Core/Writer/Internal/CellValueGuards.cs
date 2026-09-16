using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class CellValueGuards
    {
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

        internal static bool IsAlwaysFinite<T>()
        {
            return typeof(T) == typeof(int) || typeof(T) == typeof(long) || typeof(T) == typeof(short)
                || typeof(T) == typeof(byte) || typeof(T) == typeof(uint) || typeof(T) == typeof(ulong)
                || typeof(T) == typeof(ushort) || typeof(T) == typeof(sbyte) || typeof(T) == typeof(nint)
                || typeof(T) == typeof(nuint) || typeof(T) == typeof(decimal)
                || typeof(T) == typeof(Int128) || typeof(T) == typeof(UInt128);
        }

        internal static void ThrowIfNotFiniteNumberText(ReadOnlySpan<byte> utf8, Type sourceType, string paramName)
        {
            if (!double.TryParse(utf8, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                || !double.IsFinite(parsed))
            {
                ThrowUnrepresentable(sourceType, paramName);
            }
        }

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
