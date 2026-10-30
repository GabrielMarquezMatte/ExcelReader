using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using ExcelReader.Core.Reader;
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
