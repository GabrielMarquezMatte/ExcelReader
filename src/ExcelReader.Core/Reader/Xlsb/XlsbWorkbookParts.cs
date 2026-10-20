using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Core.Reader.Xlsb
{
    internal static class XlsbWorkbookParts
    {
        internal static (string Name, string Path, ExcelSheetVisibility Visibility)[] ParseSheets(ReadOnlySpan<byte> workbookBin, ReadOnlySpan<byte> relsBytes)
        {
            Dictionary<string, string> rels = XlsxXml.ParseRelationships(relsBytes);
            List<(string, string, ExcelSheetVisibility)> sheets = [];
            var reader = new Biff12RecordReader(workbookBin);
            var alternateLookup = rels.GetAlternateLookup<ReadOnlySpan<char>>();
            while (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
            {
                if (id != Brt.BundleSh)
                {
                    continue;
                }
                AddSheet(payload, alternateLookup, sheets);
            }
            return [.. sheets];
        }

        private static void AddSheet(ReadOnlySpan<byte> payload, Dictionary<string, string>.AlternateLookup<ReadOnlySpan<char>> rels, List<(string, string, ExcelSheetVisibility)> sheets)
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
            if (!rels.TryGetValue(relId, out string? target))
            {
                throw new InvalidDataException($"The workbook's sheet '{name}' refers to relationship '{relId}', which xl/_rels/workbook.bin.rels does not define.");
            }
            sheets.Add((new string(name), XlsxXml.NormalizePart(target), Visibility(Biff12.ReadU32(payload, 0))));
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
