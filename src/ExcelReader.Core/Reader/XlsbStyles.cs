namespace ExcelReader.Core.Reader
{
    // Builds the cellXfs-index -> isDate table from xl/styles.bin (binary BIFF12). A style is a date
    // when its number-format id is a builtin date format or a custom BrtFmt whose code reads as a date.
    internal static class XlsbStyles
    {
        internal static bool[] ParseStyleDateFlags(ReadOnlySpan<byte> stylesBin)
        {
            if (stylesBin.IsEmpty)
            {
                return [];
            }
            // ifmt -> isDate for custom formats; builtin ids fall back to NumberFormat.IsBuiltinDate.
            Dictionary<int, bool> custom = new(capacity: 16);
            List<bool> flags = [];
            // Only the BrtXF records between BrtBeginCellXFs/BrtEndCellXFs are the cell styles that
            // a cell's iStyleRef indexes; the earlier cellStyleXfs collection is skipped.
            bool inCellXfs = false;
            var reader = new Biff12RecordReader(stylesBin);
            while (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
            {
                switch (id)
                {
                    case Brt.Fmt:
                        ParseFmt(payload, custom);
                        break;
                    case Brt.BeginCellXFs:
                        inCellXfs = true;
                        break;
                    case Brt.EndCellXFs:
                        inCellXfs = false;
                        break;
                    case Brt.Xf when inCellXfs:
                        flags.Add(IsXfDate(payload, custom));
                        break;
                }
            }
            return [.. flags];
        }

        // BrtFmt: ifmt (u16) + stFmtCode (wide string).
        private static void ParseFmt(ReadOnlySpan<byte> payload, Dictionary<int, bool> custom)
        {
            if (payload.Length < 2)
            {
                return;
            }
            int ifmt = Biff12.ReadU16(payload, 0);
            if (Biff12.TryReadWideString(payload, 2, out ReadOnlySpan<char> code, out _))
            {
                custom[ifmt] = NumberFormat.LooksLikeDate(code);
            }
        }

        // BrtXF: ifmt (numFmtId, u16) sits at offset 2, after ixfeParent (u16).
        private static bool IsXfDate(ReadOnlySpan<byte> payload, Dictionary<int, bool> custom)
        {
            if (payload.Length < 4)
            {
                return false;
            }
            int numFmtId = Biff12.ReadU16(payload, 2);
            return custom.TryGetValue(numFmtId, out bool d) ? d : NumberFormat.IsBuiltinDate(numFmtId);
        }
    }
}
