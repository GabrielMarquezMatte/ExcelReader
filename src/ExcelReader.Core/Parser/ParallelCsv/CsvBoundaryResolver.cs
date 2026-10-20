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

        // Reads the quote state at window[start] off the first odd run of quotes there, starting from the whole of
        // a run that start cuts; even runs leave the parity alone. Under RFC 4180 a run that opens a field follows a
        // delimiter or a line start, and one that closes a field precedes a delimiter, a line end or the end of the
        // source, so a run that fits only one of the two decides. A run that fits both is null, as is one the
        // window cuts off: only the confirmed start can place that chunk. No run, or one that fits neither, is
        // guessed outside, and the confirmation after the parse still checks it.
        internal static CsvQuoteParity? InferParity(
            ReadOnlySpan<byte> window, int start, bool startsAtSource, bool endsAtEof, byte quote, byte delimiter)
        {
            int runStart = start;
            while (runStart > 0 && window[runStart - 1] == quote)
            {
                runStart--;
            }
            if (runStart == 0 && start > 0 && !startsAtSource)
            {
                return null;
            }
            bool flipped = ((start - runStart) & 1) == 1;
            int i = runStart;
            if (runStart == start)
            {
                int found = window[start..].IndexOf(quote);
                i = found < 0 ? -1 : start + found;
            }
            while (i >= 0)
            {
                int end = i + 1;
                while (end < window.Length && window[end] == quote)
                {
                    end++;
                }
                if (((end - i) & 1) == 0)
                {
                    int next = window[end..].IndexOf(quote);
                    i = next < 0 ? -1 : end + next;
                    continue;
                }
                if (end == window.Length && !endsAtEof)
                {
                    return flipped ? null : CsvQuoteParity.Outside;
                }
                if (i == 0 && !startsAtSource)
                {
                    return null;
                }
                bool closes = end == window.Length || EndsField(window[end], delimiter);
                bool opens = i == 0 || EndsField(window[i - 1], delimiter);
                if (closes && opens)
                {
                    return null;
                }
                return closes != flipped ? CsvQuoteParity.Inside : CsvQuoteParity.Outside;
            }
            return flipped ? null : CsvQuoteParity.Outside;
        }

        private static bool EndsField(byte b, byte delimiter)
        {
            return b == delimiter || b == Lf || b == Cr;
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
