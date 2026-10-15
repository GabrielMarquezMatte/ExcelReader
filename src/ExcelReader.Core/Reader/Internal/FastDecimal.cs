using System.Runtime.CompilerServices;

namespace ExcelReader.Core.Reader.Internal
{
    internal static class FastDecimal
    {
        private const int MaxMantissaDigits = 19;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static bool TryParse(ReadOnlySpan<byte> s, out decimal value)
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
            int integerStart = i;
            for (; i < s.Length; i++)
            {
                uint d = (uint)(s[i] - (byte)'0');
                if (d > 9)
                {
                    break;
                }
                mantissa = (mantissa * 10) + d;
            }
            int integerDigits = i - integerStart;
            int scale = 0;
            if (i < s.Length && s[i] == (byte)'.')
            {
                int fractionStart = ++i;
                for (; i < s.Length; i++)
                {
                    uint d = (uint)(s[i] - (byte)'0');
                    if (d > 9)
                    {
                        break;
                    }
                    mantissa = (mantissa * 10) + d;
                }
                scale = i - fractionStart;
            }
            if (i != s.Length || integerDigits + scale is 0 or > MaxMantissaDigits)
            {
                return false;
            }
            value = new decimal((int)mantissa, (int)(mantissa >> 32), 0, neg, (byte)scale);
            return true;
        }
    }
}
