using System.Runtime.CompilerServices;

namespace ExcelReader.Core.Reader.Internal
{
    internal static class FastDouble
    {
        private const int MaxExponentDigits = 3;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static bool TryParse(ReadOnlySpan<byte> s, out double value)
        {
            value = 0;
            if (s.IsEmpty)
            {
                return false;
            }
            int i = 0;
            bool neg = s[0] == (byte)'-';
            if (neg || s[0] == (byte)'+')
            {
                i = 1;
            }

            ulong mantissa = 0;
            int digits = 0;
            int scale = 0;
            int exponent = 0;
            bool sawDigit = false;
            bool sawDot = false;

            for (; i < s.Length; i++)
            {
                byte c = s[i];
                if (c == (byte)'.')
                {
                    if (sawDot)
                    {
                        return false;
                    }
                    sawDot = true;
                    continue;
                }
                if ((uint)(c - (byte)'0') > 9)
                {
                    if (c is not ((byte)'e' or (byte)'E') || !TryExponent(s[(i + 1)..], out exponent))
                    {
                        return false;
                    }
                    break;
                }
                sawDigit = true;
                bool leadingZero = mantissa == 0 && c == (byte)'0';
                if (!leadingZero)
                {
                    digits++;
                    if (digits > 15)
                    {
                        return false;
                    }
                }
                mantissa = (mantissa * 10) + (ulong)(c - (byte)'0');
                if (sawDot)
                {
                    scale++;
                }
            }
            if (!sawDigit)
            {
                return false;
            }

            scale -= exponent;
            if (scale is < -22 or > 22)
            {
                return false;
            }

            double result = mantissa;
            if (scale > 0)
            {
                result /= Pow10(scale);
            }
            else if (scale < 0)
            {
                result *= Pow10(-scale);
            }
            value = neg ? -result : result;
            return true;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static bool TryExponent(ReadOnlySpan<byte> s, out int exponent)
        {
            exponent = 0;
            if (s.IsEmpty)
            {
                return false;
            }
            int i = 0;
            bool neg = s[0] == (byte)'-';
            if (neg || s[0] == (byte)'+')
            {
                i = 1;
            }
            if (i >= s.Length || s.Length - i > MaxExponentDigits)
            {
                return false;
            }
            int value = 0;
            for (; i < s.Length; i++)
            {
                uint d = (uint)(s[i] - (byte)'0');
                if (d > 9)
                {
                    return false;
                }
                value = (value * 10) + (int)d;
            }
            exponent = neg ? -value : value;
            return true;
        }

        private static double Pow10(int scale)
        {
            return scale switch
            {
                0 => 1e0,
                1 => 1e1,
                2 => 1e2,
                3 => 1e3,
                4 => 1e4,
                5 => 1e5,
                6 => 1e6,
                7 => 1e7,
                8 => 1e8,
                9 => 1e9,
                10 => 1e10,
                11 => 1e11,
                12 => 1e12,
                13 => 1e13,
                14 => 1e14,
                15 => 1e15,
                16 => 1e16,
                17 => 1e17,
                18 => 1e18,
                19 => 1e19,
                20 => 1e20,
                21 => 1e21,
                _ => 1e22,
            };
        }
    }
}
