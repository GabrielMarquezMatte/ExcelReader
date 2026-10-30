using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer.Styles
{
    public class XlsxStyleOutputTests
    {
        private const string UnstyledStylesPart = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><numFmts count=\"1\"><numFmt numFmtId=\"14\" formatCode=\"mm-dd-yy\"/></numFmts><fonts count=\"1\"><font/></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border/></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/></cellXfs></styleSheet>";

        private static ExcelColor Color(int rgb) => ExcelColor.FromRgb(rgb);

        private static async Task<byte[]> WriteStyledCellAsync(CellStyle? style)
        {
            var ms = new MemoryStream();
            await using (var wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                int id = style is { } s ? wb.AddStyle(s) : 0;
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
                await using (XlsxRowWriter row = await sheet.StartRowAsync(id, TestContext.Current.CancellationToken))
                {
                    row.Write("x");
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }
            return ms.ToArray();
        }

        private static string StylesPart(byte[] file)
        {
            using var zip = new ZipArchive(new MemoryStream(file), ZipArchiveMode.Read);
            using var reader = new StreamReader(zip.GetEntry("xl/styles.xml")!.Open(), Encoding.UTF8);
            return reader.ReadToEnd();
        }

        private static ProbedStyle Defaults => new(null, null, null, false, false, false, false, null,
            (ExcelBorderStyle.None, null), (ExcelBorderStyle.None, null), (ExcelBorderStyle.None, null), (ExcelBorderStyle.None, null),
            ExcelHorizontalAlignment.General, ExcelVerticalAlignment.Bottom, false, 0, 0, null);

        [Fact]
        public async Task FontPropertiesRoundTrip()
        {
            byte[] file = await WriteStyledCellAsync(new CellStyle
            {
                FontName = "Consolas", FontSize = 14, FontColor = Color(0x1F4E79),
                Bold = true, Italic = true, Underline = true, Strikethrough = true,
            });
            Assert.Equal(Defaults with
            {
                FontName = "Consolas", FontSize = 14, FontRgb = 0x1F4E79,
                Bold = true, Italic = true, Underline = true, Strikethrough = true,
            }, StyleProbe.Xlsx(file, 0, 0));
        }

        [Fact]
        public async Task SolidFillRoundTrips()
        {
            byte[] file = await WriteStyledCellAsync(new CellStyle { FillColor = Color(0xFFFF00) });
            Assert.Equal(Defaults with { FillRgb = 0xFFFF00 }, StyleProbe.Xlsx(file, 0, 0));
        }

        [Fact]
        public async Task BordersRoundTripPerSide()
        {
            byte[] file = await WriteStyledCellAsync(new CellStyle
            {
                Left = new CellBorder(ExcelBorderStyle.Thin, Color(0xFF0000)),
                Right = new CellBorder(ExcelBorderStyle.Double),
                Top = new CellBorder(ExcelBorderStyle.Medium, Color(0x00FF00)),
                Bottom = new CellBorder(ExcelBorderStyle.DashDot),
            });
            Assert.Equal(Defaults with
            {
                Left = (ExcelBorderStyle.Thin, 0xFF0000),
                Right = (ExcelBorderStyle.Double, null),
                Top = (ExcelBorderStyle.Medium, 0x00FF00),
                Bottom = (ExcelBorderStyle.DashDot, null),
            }, StyleProbe.Xlsx(file, 0, 0));
        }

        [Fact]
        public async Task AlignmentRoundTrips()
        {
            byte[] file = await WriteStyledCellAsync(new CellStyle
            {
                HorizontalAlignment = ExcelHorizontalAlignment.Center, VerticalAlignment = ExcelVerticalAlignment.Top,
                WrapText = true, Indent = 3, Rotation = -45,
            });
            Assert.Equal(Defaults with
            {
                Horizontal = ExcelHorizontalAlignment.Center, Vertical = ExcelVerticalAlignment.Top,
                Wrap = true, Indent = 3, Rotation = -45,
            }, StyleProbe.Xlsx(file, 0, 0));
            Assert.Contains("textRotation=\"135\"", StylesPart(file), StringComparison.Ordinal);
        }

        [Fact]
        public async Task StackedRotationRoundTrips()
        {
            byte[] file = await WriteStyledCellAsync(new CellStyle { Rotation = CellStyle.StackedTextRotation });
            Assert.Equal(Defaults with { Rotation = 255 }, StyleProbe.Xlsx(file, 0, 0));
        }

        [Fact]
        public async Task UnstyledWorkbookStylesPartUnchanged()
        {
            byte[] file = await WriteStyledCellAsync(null);
            Assert.Equal(UnstyledStylesPart, StylesPart(file));
        }

        [Fact]
        public async Task StylesXmlValidatesInElementOrder()
        {
            byte[] file = await WriteStyledCellAsync(new CellStyle { Bold = true, FillColor = Color(1), Left = new CellBorder(ExcelBorderStyle.Thin) });
            string styles = StylesPart(file);
            int fonts = styles.IndexOf("<fonts", StringComparison.Ordinal);
            int fills = styles.IndexOf("<fills", StringComparison.Ordinal);
            int borders = styles.IndexOf("<borders", StringComparison.Ordinal);
            int cellStyleXfs = styles.IndexOf("<cellStyleXfs", StringComparison.Ordinal);
            int cellXfs = styles.IndexOf("<cellXfs", StringComparison.Ordinal);
            Assert.True(fonts < fills && fills < borders && borders < cellStyleXfs && cellStyleXfs < cellXfs);
            Assert.Contains("<fills count=\"3\">", styles, StringComparison.Ordinal);
        }
    }
}
