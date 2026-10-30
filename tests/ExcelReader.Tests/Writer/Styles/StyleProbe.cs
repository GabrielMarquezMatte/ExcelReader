using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests.Writer.Styles
{
    internal sealed record ProbedStyle(
        string? FontName, double? FontSize, int? FontRgb, bool Bold, bool Italic, bool Underline, bool Strikethrough,
        int? FillRgb, (ExcelBorderStyle Style, int? Rgb) Left, (ExcelBorderStyle Style, int? Rgb) Right,
        (ExcelBorderStyle Style, int? Rgb) Top, (ExcelBorderStyle Style, int? Rgb) Bottom,
        ExcelHorizontalAlignment Horizontal, ExcelVerticalAlignment Vertical, bool Wrap, int Indent, int Rotation,
        string? NumberFormat);

    internal static class StyleProbe
    {
        private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        internal static ProbedStyle Xlsx(byte[] file, int row, int column)
        {
            int styleIndex = StyleIndexOf(Excel.FromXlsx(new MemoryStream(file)), row, column);
            using var zip = new ZipArchive(new MemoryStream(file), ZipArchiveMode.Read);
            XElement sheet = XDocument.Load(zip.GetEntry("xl/styles.xml")!.Open()).Root!;
            XElement xf = sheet.Element(Main + "cellXfs")!.Elements(Main + "xf").ElementAt(styleIndex);
            XElement font = Nth(sheet, "fonts", "font", xf, "fontId");
            XElement fill = Nth(sheet, "fills", "fill", xf, "fillId");
            XElement border = Nth(sheet, "borders", "border", xf, "borderId");
            XElement? alignment = xf.Element(Main + "alignment");
            return new ProbedStyle(
                font.Element(Main + "name")?.Attribute("val")?.Value,
                font.Element(Main + "sz") is { } sz ? double.Parse(sz.Attribute("val")!.Value, CultureInfo.InvariantCulture) : null,
                Rgb(font.Element(Main + "color")),
                font.Element(Main + "b") is not null,
                font.Element(Main + "i") is not null,
                font.Element(Main + "u") is not null,
                font.Element(Main + "strike") is not null,
                Rgb(fill.Element(Main + "patternFill")?.Element(Main + "fgColor")),
                Side(border, "left"), Side(border, "right"), Side(border, "top"), Side(border, "bottom"),
                Enum.Parse<ExcelHorizontalAlignment>(alignment?.Attribute("horizontal")?.Value ?? "general", ignoreCase: true),
                Enum.Parse<ExcelVerticalAlignment>(alignment?.Attribute("vertical")?.Value ?? "bottom", ignoreCase: true),
                string.Equals(alignment?.Attribute("wrapText")?.Value, "1", StringComparison.Ordinal),
                int.Parse(alignment?.Attribute("indent")?.Value ?? "0", CultureInfo.InvariantCulture),
                RotationOf(int.Parse(alignment?.Attribute("textRotation")?.Value ?? "0", CultureInfo.InvariantCulture)),
                NumberFormatOf(sheet, int.Parse(xf.Attribute("numFmtId")!.Value, CultureInfo.InvariantCulture)));
        }

        private static readonly ExcelVerticalAlignment[] BiffVertical =
            [ExcelVerticalAlignment.Top, ExcelVerticalAlignment.Center, ExcelVerticalAlignment.Bottom,
             ExcelVerticalAlignment.Justify, ExcelVerticalAlignment.Distributed];

        internal static ProbedStyle Xlsb(byte[] file, int row, int column)
        {
            int styleIndex = StyleIndexOf(Excel.FromXlsb(new MemoryStream(file)), row, column);
            List<(int Id, byte[] Payload)> records = XlsbRecords(file);
            byte[][] fonts = Payloads(records, Brt.Font);
            int xfStart = records.FindIndex(r => r.Id == Brt.BeginCellXFs) + 1;
            byte[] xf = records[xfStart + styleIndex].Payload;
            byte[] font = fonts[BitConverter.ToUInt16(xf, 4)];
            byte[] fill = Payloads(records, Brt.Fill)[BitConverter.ToUInt16(xf, 6)];
            byte[] border = Payloads(records, Brt.Border)[BitConverter.ToUInt16(xf, 8)];
            string name = FontName(font);
            int height = BitConverter.ToUInt16(font, 0);
            int flags = BitConverter.ToUInt16(xf, 12);
            bool defaultName = string.Equals(name, FontName(fonts[0]), StringComparison.Ordinal);
            bool defaultHeight = height == BitConverter.ToUInt16(fonts[0], 0);
            return new ProbedStyle(
                defaultName ? null : name,
                defaultHeight ? null : height / 20.0,
                BiffRgb(font, 12),
                BitConverter.ToUInt16(font, 4) == 700,
                (font[2] & 0x02) != 0,
                font[8] != 0,
                (font[2] & 0x08) != 0,
                BiffRgb(fill, 4),
                BiffSide(border, 21), BiffSide(border, 31), BiffSide(border, 1), BiffSide(border, 11),
                (ExcelHorizontalAlignment)(flags & 0x7),
                BiffVertical[(flags >> 3) & 0x7],
                (flags & 0x40) != 0,
                xf[11],
                RotationOf(xf[10]),
                BiffNumberFormat(records, BitConverter.ToUInt16(xf, 2)));
        }

        internal static byte[] XlsbStylesPart(byte[] file)
        {
            using var zip = new ZipArchive(new MemoryStream(file), ZipArchiveMode.Read);
            using var ms = new MemoryStream();
            zip.GetEntry("xl/styles.bin")!.Open().CopyTo(ms);
            return ms.ToArray();
        }

        internal static List<(int Id, byte[] Payload)> XlsbRecords(byte[] file)
        {
            byte[] data = XlsbStylesPart(file);
            var records = new List<(int, byte[])>();
            int pos = 0;
            while (pos < data.Length)
            {
                int id = Varint(data, ref pos, 2);
                int length = Varint(data, ref pos, 4);
                records.Add((id, data[pos..(pos + length)]));
                pos += length;
            }
            return records;
        }

        private static int Varint(byte[] data, ref int pos, int maxBytes)
        {
            int value = 0;
            for (int i = 0; i < maxBytes; i++)
            {
                byte b = data[pos++];
                value |= (b & 0x7F) << (7 * i);
                if ((b & 0x80) == 0)
                {
                    break;
                }
            }
            return value;
        }

        private static byte[][] Payloads(List<(int Id, byte[] Payload)> records, int id)
        {
            return records.Where(r => r.Id == id).Select(r => r.Payload).ToArray();
        }

        private static string FontName(byte[] font)
        {
            return Encoding.Unicode.GetString(font, 25, BitConverter.ToInt32(font, 21) * 2);
        }

        // BrtColor: xColorType 2 (RGB) in bits 1-7 of the first byte, then R, G, B at offsets 4-6.
        private static int? BiffRgb(byte[] payload, int offset)
        {
            if (payload[offset] >> 1 != 2)
            {
                return null;
            }
            return (payload[offset + 4] << 16) | (payload[offset + 5] << 8) | payload[offset + 6];
        }

        private static (ExcelBorderStyle Style, int? Rgb) BiffSide(byte[] border, int offset)
        {
            var style = (ExcelBorderStyle)border[offset];
            if (style == ExcelBorderStyle.None)
            {
                return (ExcelBorderStyle.None, null);
            }
            return (style, BiffRgb(border, offset + 2));
        }

        private static string? BiffNumberFormat(List<(int Id, byte[] Payload)> records, int numFmtId)
        {
            if (numFmtId == 0)
            {
                return null;
            }
            byte[]? fmt = Payloads(records, Brt.Fmt).FirstOrDefault(p => BitConverter.ToUInt16(p, 0) == numFmtId);
            if (fmt is null)
            {
                return null;
            }
            return Encoding.Unicode.GetString(fmt, 6, BitConverter.ToInt32(fmt, 2) * 2);
        }

        private static int StyleIndexOf(IExcelWorkbook reader, int row, int column)
        {
            using (reader)
            {
                using var e = reader.FirstSheet.GetEnumerator();
                for (int i = 0; i <= row; i++)
                {
                    Assert.True(e.MoveNext());
                }
                return e.Current[column].StyleIndex;
            }
        }

        private static XElement Nth(XElement sheet, string list, string item, XElement xf, string idAttribute)
        {
            int id = int.Parse(xf.Attribute(idAttribute)!.Value, CultureInfo.InvariantCulture);
            return sheet.Element(Main + list)!.Elements(Main + item).ElementAt(id);
        }

        private static int? Rgb(XElement? color)
        {
            string? rgb = color?.Attribute("rgb")?.Value;
            return rgb is null ? null : int.Parse(rgb, NumberStyles.HexNumber, CultureInfo.InvariantCulture) & 0xFFFFFF;
        }

        private static (ExcelBorderStyle Style, int? Rgb) Side(XElement border, string name)
        {
            XElement? side = border.Element(Main + name);
            string? style = side?.Attribute("style")?.Value;
            if (style is null)
            {
                return (ExcelBorderStyle.None, null);
            }
            return (Enum.Parse<ExcelBorderStyle>(style, ignoreCase: true), Rgb(side!.Element(Main + "color")));
        }

        private static int RotationOf(int stored)
        {
            if (stored > 90 && stored <= 180)
            {
                return 90 - stored;
            }
            return stored;
        }

        private static string? NumberFormatOf(XElement sheet, int numFmtId)
        {
            if (numFmtId == 0)
            {
                return null;
            }
            return sheet.Element(Main + "numFmts")!.Elements(Main + "numFmt")
                .FirstOrDefault(n => string.Equals(n.Attribute("numFmtId")!.Value, numFmtId.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal))?
                .Attribute("formatCode")?.Value;
        }
    }
}
