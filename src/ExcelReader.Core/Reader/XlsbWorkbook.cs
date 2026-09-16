namespace ExcelReader.Core.Reader
{
    internal static class XlsbWorkbook
    {
        internal static (string Name, string Path)[] ParseSheets(ReadOnlySpan<byte> workbookBin, ReadOnlySpan<byte> relsBytes)
        {
            Dictionary<string, string> rels = XlsxXml.ParseRelationships(relsBytes);
            List<(string, string)> sheets = [];
            var reader = new Biff12RecordReader(workbookBin);
            while (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
            {
                if (id != Brt.BundleSh)
                {
                    continue;
                }
                AddSheet(payload, rels, sheets);
            }
            return [.. sheets];
        }

        private static void AddSheet(ReadOnlySpan<byte> payload, Dictionary<string, string> rels, List<(string, string)> sheets)
        {
            if (payload.Length < 8)
            {
                return;
            }
            if (!Biff12.TryReadWideString(payload, 8, out ReadOnlySpan<char> relId, out int consumed))
            {
                return;
            }
            if (!Biff12.TryReadWideString(payload, 8 + consumed, out ReadOnlySpan<char> name, out _))
            {
                return;
            }
            if (rels.TryGetValue(new string(relId), out string? target))
            {
                sheets.Add((new string(name), XlsxXml.NormalizePart(target)));
            }
        }

        internal static bool ParseDate1904(ReadOnlySpan<byte> workbookBin)
        {
            var reader = new Biff12RecordReader(workbookBin);
            while (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
            {
                if (id == Brt.WbProp && payload.Length >= 4)
                {
                    return (Biff12.ReadU32(payload, 0) & 0x01) != 0;
                }
            }
            return false;
        }
    }
}
