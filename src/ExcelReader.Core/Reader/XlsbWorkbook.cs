using ExcelReader.Core.Enums;

namespace ExcelReader.Core.Reader
{
    internal static class XlsbWorkbook
    {
        internal static (string Name, string Path, ExcelSheetVisibility Visibility)[] ParseSheets(ReadOnlySpan<byte> workbookBin, ReadOnlySpan<byte> relsBytes)
        {
            Dictionary<string, string> rels = XlsxXml.ParseRelationships(relsBytes);
            List<(string, string, ExcelSheetVisibility)> sheets = [];
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

        private static void AddSheet(ReadOnlySpan<byte> payload, Dictionary<string, string> rels, List<(string, string, ExcelSheetVisibility)> sheets)
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
                sheets.Add((new string(name), XlsxXml.NormalizePart(target), Visibility(Biff12.ReadU32(payload, 0))));
            }
        }

        private static ExcelSheetVisibility Visibility(uint state)
        {
            return state switch
            {
                1 => ExcelSheetVisibility.Hidden,
                2 => ExcelSheetVisibility.VeryHidden,
                _ => ExcelSheetVisibility.Visible,
            };
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
