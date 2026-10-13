namespace ExcelReader.Core.Writer.Xls
{
    /// <summary>
    /// Encodes a number as an RkNumber, the same in [MS-XLS] and [MS-XLSB]: a 30-bit integer, or the top
    /// 30 bits of a double, either one optionally divided by 100. A value is encoded only when decoding
    /// gives back the same bits, so RK never changes a number; anything else keeps an 8-byte double.
    /// </summary>
    internal static class XlsRk
    {
        private const double IntLimit = 1 << 29;
        private const long LowBits = 0x3_FFFF_FFFF;

        internal static bool TryEncode(double value, out int rk)
        {
            // Excel picks the truncated double whenever it is exact, so that form goes first.
            long bits = BitConverter.DoubleToInt64Bits(value);
            if ((bits & LowBits) == 0)
            {
                rk = (int)(bits >> 32);
                return true;
            }
            if (Math.Abs(value) < IntLimit)
            {
                int whole = (int)value;
                if (whole == value)
                {
                    rk = (whole << 2) | 2;
                    return true;
                }
            }

            // Only these two forms divide by 100 on decoding, so only they need decoding to prove exact.
            // value * 100 can miss an integer by an ulp (0.07 * 100) while n / 100 still gives value back.
            double scaled = value * 100;
            if (Math.Abs(scaled) < IntLimit)
            {
                double hundredths = Math.Round(scaled);
                if (Math.Abs(scaled - hundredths) < 1e-6 && Matches((((int)hundredths) << 2) | 3, bits, out rk))
                {
                    return true;
                }
            }
            long scaledBits = BitConverter.DoubleToInt64Bits(scaled);
            if ((scaledBits & LowBits) == 0 && Matches(((int)(scaledBits >> 32) & ~3) | 1, bits, out rk))
            {
                return true;
            }
            rk = 0;
            return false;
        }

        internal static double Decode(int rk)
        {
            double value = (rk & 0x02) != 0
                ? rk >> 2
                : BitConverter.Int64BitsToDouble((long)((ulong)(uint)(rk & ~3) << 32));
            return (rk & 0x01) != 0 ? value / 100.0 : value;
        }

        private static bool Matches(int candidate, long bits, out int rk)
        {
            rk = candidate;
            return BitConverter.DoubleToInt64Bits(Decode(candidate)) == bits;
        }
    }
}
