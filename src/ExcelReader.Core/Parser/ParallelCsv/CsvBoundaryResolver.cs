namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal enum CsvQuoteParity : byte
    {
        Outside = 0,
        Inside = 1,
    }

    internal static class CsvBoundaryResolver
    {
        private const byte Cr = (byte)'\r';
        private const byte Lf = (byte)'\n';

        internal static bool StartsRecord(ReadOnlySpan<byte> pair)
        {
            if (pair[0] == Lf)
            {
                return true;
            }
            return pair[0] == Cr && pair[1] != Lf;
        }

        internal static int FindRecordStart(ReadOnlySpan<byte> window, byte quote, CsvQuoteParity parity)
        {
            bool inside = parity == CsvQuoteParity.Inside;
            for (int i = 0; i < window.Length; i++)
            {
                byte b = window[i];
                if (b == quote)
                {
                    inside = !inside;
                    continue;
                }
                if (inside)
                {
                    continue;
                }
                if (b == Lf)
                {
                    return NextOrMinusOne(window, i + 1);
                }
                if (b != Cr)
                {
                    continue;
                }
                int afterCr = i + 1 < window.Length && window[i + 1] == Lf ? i + 2 : i + 1;
                return NextOrMinusOne(window, afterCr);
            }
            return -1;
        }

        private static int NextOrMinusOne(ReadOnlySpan<byte> window, int index)
        {
            return index < window.Length ? index : -1;
        }
    }
}
