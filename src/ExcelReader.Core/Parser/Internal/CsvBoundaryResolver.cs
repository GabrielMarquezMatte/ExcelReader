namespace ExcelReader.Core.Parser.Internal
{
    // Which side of a quoted field a scan starts on. Inside a quoted field a `"` is either the
    // closing quote or the first half of an escaped `""`; either way it toggles the state, and a
    // doubled quote toggles twice and lands back inside. So "am I inside a quoted field" is exactly
    // the parity of the quote count, and the two boundary hypotheses are just the two parities.
    internal enum CsvQuoteParity : byte
    {
        Outside = 0,
        Inside = 1,
    }

    // Locates the first record start in a window of bytes read from an arbitrary file offset, under
    // an assumed quote parity. A chunk cannot know its own parity locally — that is what the
    // predecessor chunk's ResolvedNextStart confirms during the ordered merge.
    internal static class CsvBoundaryResolver
    {
        private const byte Cr = (byte)'\r';
        private const byte Lf = (byte)'\n';

        // Returns the index within `window` of the first byte beginning a record, or -1 when the
        // window contains no usable boundary. Terminator handling mirrors the record parser
        // (CsvReader.Enumerator.TryParseSimpleRecord): \n ends a record, a lone \r ends a record,
        // and \r\n counts as a single terminator.
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
                // A \r as the window's final byte needs no special case: whether or not a \n follows
                // it, the record starts at or past the window's end, and NextOrMinusOne rejects both.
                // The caller reads one byte beyond the chunk precisely so a \r at the chunk's last
                // position is resolved against real data rather than a guess.
                int afterCr = i + 1 < window.Length && window[i + 1] == Lf ? i + 2 : i + 1;
                return NextOrMinusOne(window, afterCr);
            }
            return -1;
        }

        // A boundary whose record start lands exactly at the end of the window is not a record start
        // *within* this window — there are no bytes there to parse.
        private static int NextOrMinusOne(ReadOnlySpan<byte> window, int index)
        {
            return index < window.Length ? index : -1;
        }
    }
}
