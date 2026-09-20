namespace ExcelReader.Core.Reader
{
    internal static class XlsbStyles
    {
        internal static bool[] ParseStyleDateFlags(ReadOnlySpan<byte> stylesBin)
        {
            if (stylesBin.IsEmpty)
            {
                return [];
            }
            Dictionary<int, bool> custom = new(capacity: 16);
            List<bool> flags = [];
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

        private static bool IsXfDate(ReadOnlySpan<byte> payload, Dictionary<int, bool> custom)
        {
            if (payload.Length < 4)
            {
                return false;
            }
            int numFmtId = Biff12.ReadU16(payload, 2);
            return WorkbookLookups.ResolveDateFlag(custom, numFmtId);
        }
    }
}
