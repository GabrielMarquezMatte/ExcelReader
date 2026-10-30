using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsb;

namespace ExcelReader.Tests.Writer.Styles
{
    public class XlsbStyleOutputTests
    {
        private const string UnstyledStylesPart = "960200E7040400000000E80400E30404010000002B27DC000000900100000002000000000000000000000007000000430061006C006900620072006900E40400DB0404020000002D440000000003400000000000FF03410000FFFFFFFF0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000002D441100000003400000000000FF03410000FFFFFFFF000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000DC0400E50404010000002E33000000010000000000000000000100000000000000000001000000000000000000010000000000000000000100000000000000E60400F20404010000002F10FFFF0000000000000000000010100000F30400E90404020000002F10000000000000000000000000101000002F1000000E00000000000000000010100100EA0400EB04040100000030180000000001000000060000004E006F0072006D0061006C00EC0400970200";

        private static ExcelColor Color(int rgb)
        {
            return ExcelColor.FromRgb(rgb);
        }

        private static async Task<byte[]> WriteStyledCellAsync(CellStyle? style)
        {
            var ms = new MemoryStream();
            await using (var wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                int id = style is { } s ? wb.AddStyle(s) : 0;
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                await using (XlsbRowWriter row = await sheet.StartRowAsync(id, TestContext.Current.CancellationToken))
                {
                    row.Write("x");
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }
            return ms.ToArray();
        }

        private static ProbedStyle Defaults => new(null, null, null, false, false, false, false, null,
            (ExcelBorderStyle.None, null), (ExcelBorderStyle.None, null), (ExcelBorderStyle.None, null), (ExcelBorderStyle.None, null),
            ExcelHorizontalAlignment.General, ExcelVerticalAlignment.Bottom, false, 0, 0, null);

        [Fact]
        public async Task XlsbFontPropertiesRoundTrip()
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
            }, StyleProbe.Xlsb(file, 0, 0));
        }

        [Fact]
        public async Task XlsbSolidFillRoundTrips()
        {
            byte[] file = await WriteStyledCellAsync(new CellStyle { FillColor = Color(0xFFFF00) });
            Assert.Equal(Defaults with { FillRgb = 0xFFFF00 }, StyleProbe.Xlsb(file, 0, 0));
        }

        [Fact]
        public async Task XlsbBordersRoundTripPerSide()
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
            }, StyleProbe.Xlsb(file, 0, 0));
        }

        [Fact]
        public async Task XlsbAlignmentRoundTrips()
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
            }, StyleProbe.Xlsb(file, 0, 0));
            byte[] xf = StyleProbe.XlsbRecords(file).Where(r => r.Id == Brt.Xf).Last().Payload;
            Assert.Equal(135, xf[10]);
        }

        [Fact]
        public async Task XlsbStackedRotationRoundTrips()
        {
            byte[] file = await WriteStyledCellAsync(new CellStyle { Rotation = CellStyle.StackedTextRotation });
            Assert.Equal(Defaults with { Rotation = 255 }, StyleProbe.Xlsb(file, 0, 0));
        }

        [Fact]
        public async Task XlsbUnstyledWorkbookStylesPartUnchanged()
        {
            byte[] file = await WriteStyledCellAsync(null);
            Assert.Equal(UnstyledStylesPart, Convert.ToHexString(StyleProbe.XlsbStylesPart(file)));
        }

        [Fact]
        public async Task XlsbStylesRecordsInOrder()
        {
            byte[] file = await WriteStyledCellAsync(new CellStyle { Bold = true, FillColor = Color(1), Left = new CellBorder(ExcelBorderStyle.Thin) });
            List<int> ids = StyleProbe.XlsbRecords(file).Select(r => r.Id).ToList();
            int fonts = ids.IndexOf(Brt.BeginFonts);
            int fills = ids.IndexOf(Brt.BeginFills);
            int borders = ids.IndexOf(Brt.BeginBorders);
            int cellStyleXfs = ids.IndexOf(Brt.BeginCellStyleXFs);
            int cellXfs = ids.IndexOf(Brt.BeginCellXFs);
            Assert.True(fonts < fills && fills < borders && borders < cellStyleXfs && cellStyleXfs < cellXfs);
            byte[] fillCount = StyleProbe.XlsbRecords(file).First(r => r.Id == Brt.BeginFills).Payload;
            Assert.Equal(3, BitConverter.ToInt32(fillCount));
        }

        [Fact]
        public async Task BoldFromBeforeThisChangeIsNowWritten()
        {
            byte[] file = await WriteStyledCellAsync(new CellStyle { Bold = true });
            Assert.Equal(Defaults with { Bold = true }, StyleProbe.Xlsb(file, 0, 0));
        }
    }
}
