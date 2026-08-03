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
                sawDigit = true;
                // A leading zero (mantissa still 0 and this digit is itself 0) contributes nothing to
                // the value and isn't a significant digit, so it doesn't count against the 15-digit
                // cap — only against `scale` if it happens to fall after the decimal point, which it
                // still must (0.0000001 needs every one of those zeros to compute the right magnitude).
                // Without this, a value like "000000000000000123" would hit the cap on padding alone
                // and fall back to the general parser despite having only 3 significant digits.
                bool leadingZero = mantissa == 0 && c == (byte)'0';
                if (!leadingZero)
                {
                    digits++;
                    if (digits > 15)
                    {
                        return false; // mantissa would no longer fit a double's exact integer range
                    }
                }
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
                result /= Pow10(scale);
            }
            value = neg ? -result : result;
            return true;
        }

        // Every power of ten up to 1e22 is itself exactly representable as a double (see the class
        // remarks); a switch lets the JIT emit a jump table of immediate constants instead of a static
        // array load, skipping both the static-cctor-check and the bounds check an array indexer pays.
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
