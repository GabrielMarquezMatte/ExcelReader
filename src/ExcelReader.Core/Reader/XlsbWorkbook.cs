using System.Buffers;
using System.Text;

namespace ExcelReader.Core.Reader
{
    // Parses xl/workbook.bin (binary BIFF12): the sheet bundle and the date system. Sheet names come
    // from BrtBundleSh records; their part paths come from the (still XML) workbook.bin.rels.
    internal static class XlsbWorkbook
    {
        internal static (string Name, string Path)[] ParseSheets(ReadOnlySpan<byte> workbookBin, ReadOnlySpan<byte> relsBytes)
        {
            Dictionary<string, string> rels = ParseRels(relsBytes);
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

        // BrtBundleSh: hsState (u32), iTabID (u32), strRelID (nullable wide string), strName (wide string).
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
                sheets.Add((new string(name), NormalizePart(target)));
            }
        }

        // BrtWbProp flags (u32): bit 0 (f1904) selects the 1904 date system.
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

        private static Dictionary<string, string> ParseRels(ReadOnlySpan<byte> relsBytes)
        {
            Dictionary<string, string> rels = new(StringComparer.Ordinal);
            if (relsBytes.IsEmpty)
            {
                return rels;
            }
            foreach (ReadOnlySpan<byte> tag in new TagSpanEnumerable(relsBytes, "<Relationship"u8))
            {
                string id = DecodeAttr(XlsxXml.Attr(tag, " Id=\""u8));
                if (id.Length > 0)
                {
                    rels[id] = DecodeAttr(XlsxXml.Attr(tag, " Target=\""u8));
                }
            }
            return rels;
        }

        private static string NormalizePart(ReadOnlySpan<char> target)
        {
            if (target.Length > 0 && target[0] == '/')
            {
                return new string(target[1..]);
            }
            if (target.StartsWith("xl/"))
            {
                return new string(target);
            }
            return $"xl/{target}";
        }

        private static string DecodeAttr(ReadOnlySpan<byte> raw)
        {
            if (raw.IsEmpty)
            {
                return string.Empty;
            }
            byte[] buffer = ArrayPool<byte>.Shared.Rent(raw.Length);
            try
            {
                int written = XlsxXml.Decode(raw, buffer);
                return Encoding.UTF8.GetString(buffer.AsSpan(0, written));
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }
    }
}
