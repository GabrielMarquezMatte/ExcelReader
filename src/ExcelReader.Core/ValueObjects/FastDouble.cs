namespace ExcelReader.Core.ValueObjects
{
    // Parses a plain (non-exponent) ASCII decimal number straight to a double without the general
    // culture-aware double.TryParse machinery. Succeeds only when the value is exactly representable
    // via `mantissa / 10^scale`: mantissa fits in <= 15 decimal digits (well under a double's 53-bit
    // integer precision) and scale is <= 22 (every power of ten up to 1e22 is itself exactly
    // representable as a double). Under those bounds, IEEE 754 division is correctly rounded, so the
    // result is bit-identical to double.TryParse(InvariantCulture) for every input this accepts.
    // Anything else (exponents, too many digits, malformed text) returns false so the caller can fall
    // back to the general parser.
    internal static class FastDouble
    {
        private static readonly double[] Pow10 =
        [
            1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10, 1e11,
            1e12, 1e13, 1e14, 1e15, 1e16, 1e17, 1e18, 1e19, 1e20, 1e21, 1e22,
        ];

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
                    return false; // exponent marker or anything else not handled here
                }
                digits++;
                if (digits > 15)
                {
                    return false; // mantissa would no longer fit a double's exact integer range
                }
                sawDigit = true;
                mantissa = (mantissa * 10) + (ulong)(c - (byte)'0');
                if (sawDot)
                {
                    scale++;
                }
            }
            if (!sawDigit || scale > 22)
            {
                return false;
            }

            double result = mantissa;
            if (scale > 0)
            {
                result /= Pow10[scale];
            }
            value = neg ? -result : result;
            return true;
        }
    }
}
