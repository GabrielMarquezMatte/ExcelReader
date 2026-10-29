using System.Buffers.Text;

namespace ExcelReader.Core.Writer.Internal
{
    /// <summary>
    /// Shortest round-trip digits for 1e-3 &lt;= |v| &lt; 1e15, byte-identical to <see cref="Utf8Formatter"/>.
    /// v·10^k is held exactly as N / 2^s. Its rounding interval is symmetric (mantissa not a power of two)
    /// and its ends are never integers in this range (2^s would have to divide 10^k), so the shortest form
    /// is the multiple of the largest power of ten strictly inside it nearest to v. Exact ties go to
    /// <see cref="Utf8Formatter"/>.
    /// </summary>
    internal static class ShortestDouble
    {
        private static ReadOnlySpan<ulong> Pow10 =>
        [
            1, 10, 100, 1_000, 10_000, 100_000, 1_000_000, 10_000_000, 100_000_000, 1_000_000_000,
            10_000_000_000, 100_000_000_000, 1_000_000_000_000, 10_000_000_000_000, 100_000_000_000_000,
            1_000_000_000_000_000, 10_000_000_000_000_000, 100_000_000_000_000_000,
            1_000_000_000_000_000_000, 10_000_000_000_000_000_000,
        ];

        // False when the value is out of range, an exact tie, or does not fit; Utf8Formatter decides then.
        internal static bool TryFormat(double value, Span<byte> destination, out int written)
        {
            written = 0;
            double abs = Math.Abs(value);
            ulong bits = BitConverter.DoubleToUInt64Bits(abs);
            ulong fraction = bits & ((1UL << 52) - 1);
            if (!(abs >= 1e-3 && abs < 1e15) || fraction == 0)
            {
                return false;
            }
            int exponent = (int)(bits >> 52) - 1075;
            int shift = 1 - exponent;
            int k = 16 - (((exponent + 52) * 1233) >> 12);
            UInt128 scaled = Scale((fraction | (1UL << 52)) << 1, k);
            UInt128 halfUlp = Scale(1, k);
            ulong below = (ulong)((scaled - halfUlp) >> shift);
            ulong inside = (ulong)((scaled + halfUlp) >> shift);
            int dropped = 0;
            while (inside / 10 > below / 10)
            {
                inside /= 10;
                below /= 10;
                dropped++;
            }

            ulong integer = (ulong)(scaled >> shift);
            UInt128 rest = scaled & ((UInt128.One << shift) - 1);
            ulong unit = Pow10[dropped];
            ulong digits = integer / unit;
            UInt128 twiceRemainder = ((UInt128)(integer - (digits * unit)) << (shift + 1)) + (rest << 1);
            UInt128 unitScaled = (UInt128)unit << shift;
            if (twiceRemainder == unitScaled)
            {
                return false;
            }
            if (twiceRemainder > unitScaled)
            {
                digits++;
            }
            int exp10 = dropped - k;
            while (digits % 10 == 0)
            {
                digits /= 10;
                exp10++;
            }
            return Write(value < 0, digits, exp10, destination, out written);
        }

        private static UInt128 Scale(ulong value, int k)
        {
            if (k > 19)
            {
                value *= 10;
                k = 19;
            }
            ulong high = Math.BigMul(value, Pow10[k], out ulong low);
            return new UInt128(high, low);
        }

        private static bool Write(bool negative, ulong digits, int exp10, Span<byte> destination, out int written)
        {
            written = 0;
            Span<byte> text = stackalloc byte[20];
            Utf8Formatter.TryFormat(digits, text, out int count);
            int point = count + exp10;
            if (destination.Length < (negative ? 1 : 0) + TextLength(count, exp10, point))
            {
                return false;
            }
            if (negative)
            {
                destination[written++] = (byte)'-';
            }
            if (exp10 >= 0)
            {
                text[..count].CopyTo(destination[written..]);
                destination.Slice(written + count, exp10).Fill((byte)'0');
                written += point;
                return true;
            }
            if (point > 0)
            {
                text[..point].CopyTo(destination[written..]);
                destination[written + point] = (byte)'.';
                text[point..count].CopyTo(destination[(written + point + 1)..]);
                written += count + 1;
                return true;
            }
            destination[written] = (byte)'0';
            destination[written + 1] = (byte)'.';
            destination.Slice(written + 2, -point).Fill((byte)'0');
            text[..count].CopyTo(destination[(written + 2 - point)..]);
            written += 2 - point + count;
            return true;
        }

        private static int TextLength(int count, int exp10, int point)
        {
            if (exp10 >= 0)
            {
                return point;
            }
            if (point > 0)
            {
                return count + 1;
            }
            return 2 - point + count;
        }
    }
}
