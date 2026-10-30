using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xls;

namespace ExcelReader.Tests.Writer.Styles
{
    public class XlsStyleOutputTests
    {
        private const string UnstyledGlobals = "080900060500BB0DCC07000000000600000000E1B00400C1000000E2005C0B0000457863656C52656164657220202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020202020200042B00401610000013D0100009C110000190000001200000013000001AF000001BC0000003DFF7FFF7F807030393800000000000100580200400000008D000000220000000E010001B7000000DA00000031DC000000FF7F90010000000000000500417269616C0031DC000000FF7F90010000000000000500417269616C0031DC000000FF7F90010000000000000500417269616C0031DC000000FF7F90010000000000000500417269616C00E000000000F5FF200000000000000000000000C02000E000000000F5FF200000000000000000000000C02000E000000000F5FF200000000000000000000000C02000E000000000F5FF200000000000000000000000C02000E000000000F5FF200000000000000000000000C02000E000000000F5FF200000000000000000000000C02000E000000000F5FF200000000000000000000000C02000E000000000F5FF200000000000000000000000C02000E000000000F5FF200000000000000000000000C02000E000000000F5FF200000000000000000000000C02000E000000000F5FF200000000000000000000000C02000E000000000F5FF200000000000000000000000C02000E000000000F5FF200000000000000000000000C02000E000000000F5FF200000000000000000000000C02000E000000000F5FF200000000000000000000000C02000E000000000F5FF200000000000000000000000C02000E0000000000100200000000000000000000002C02000E000000E000100200000000000000000000002C0200293008000FF00853A03000000000600536865657431000A";

        private const int BoundSheet = 0x0085;
        private const int Eof = 0x000A;

        private static ExcelColor Color(int rgb)
        {
            return ExcelColor.FromRgb(rgb);
        }

        private static byte[] WriteStyledCells(params CellStyle[] styles)
        {
            var ms = new MemoryStream();
            using (var wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                int[] ids = styles.Select(wb.AddStyle).ToArray();
                using XlsSheetWriter sheet = wb.AddSheet("Sheet1");
                foreach (int id in ids)
                {
                    using XlsRowWriter row = sheet.StartRow(id);
                    row.Write("x");
                }
                sheet.End();
                wb.End();
            }
            return ms.ToArray();
        }

        private static byte[] WriteUnstyledCell()
        {
            var ms = new MemoryStream();
            using (var wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                using XlsSheetWriter sheet = wb.AddSheet("Sheet1");
                using (XlsRowWriter row = sheet.StartRow())
                {
                    row.Write("x");
                }
                sheet.End();
                wb.End();
            }
            return ms.ToArray();
        }

        private static List<XlsRecord> Globals(byte[] file)
        {
            List<XlsRecord> records = StyleProbe.XlsRecords(file);
            return records.GetRange(0, records.FindIndex(r => r.Id == Eof) + 1);
        }

        private static ProbedStyle Defaults => new(null, null, null, false, false, false, false, null,
            (ExcelBorderStyle.None, null), (ExcelBorderStyle.None, null), (ExcelBorderStyle.None, null), (ExcelBorderStyle.None, null),
            ExcelHorizontalAlignment.General, ExcelVerticalAlignment.Bottom, false, 0, 0, null);

        [Fact]
        public void XlsFontPropertiesRoundTrip()
        {
            byte[] file = WriteStyledCells(new CellStyle
            {
                FontName = "Consolas", FontSize = 14, FontColor = Color(0x1F4E79),
                Bold = true, Italic = true, Underline = true, Strikethrough = true,
            });
            Assert.Equal(Defaults with
            {
                FontName = "Consolas", FontSize = 14, FontRgb = 0x1F4E79,
                Bold = true, Italic = true, Underline = true, Strikethrough = true,
            }, StyleProbe.Xls(file, 0, 0));
        }

        [Fact]
        public void XlsSolidFillRoundTrips()
        {
            byte[] file = WriteStyledCells(new CellStyle { FillColor = Color(0xFFFF00) });
            Assert.Equal(Defaults with { FillRgb = 0xFFFF00 }, StyleProbe.Xls(file, 0, 0));
        }

        [Fact]
        public void XlsBordersRoundTripPerSide()
        {
            byte[] file = WriteStyledCells(new CellStyle
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
            }, StyleProbe.Xls(file, 0, 0));
        }

        [Fact]
        public void XlsAlignmentRoundTrips()
        {
            byte[] file = WriteStyledCells(new CellStyle
            {
                HorizontalAlignment = ExcelHorizontalAlignment.Center, VerticalAlignment = ExcelVerticalAlignment.Top,
                WrapText = true, Indent = 3, Rotation = -45,
            });
            Assert.Equal(Defaults with
            {
                Horizontal = ExcelHorizontalAlignment.Center, Vertical = ExcelVerticalAlignment.Top,
                Wrap = true, Indent = 3, Rotation = -45,
            }, StyleProbe.Xls(file, 0, 0));
            byte[] xf = StyleProbe.XlsPayloads(StyleProbe.XlsRecords(file), StyleProbe.BiffXf)[^1];
            Assert.Equal(135, xf[7]);
        }

        [Fact]
        public void XlsStackedRotationRoundTrips()
        {
            byte[] file = WriteStyledCells(new CellStyle { Rotation = CellStyle.StackedTextRotation });
            Assert.Equal(Defaults with { Rotation = 255 }, StyleProbe.Xls(file, 0, 0));
        }

        [Fact]
        public void XlsUnstyledWorkbookGlobalsUnchanged()
        {
            byte[] file = WriteUnstyledCell();
            string globals = string.Concat(Globals(file).Select(r => $"{r.Id:X4}{Convert.ToHexString(r.Payload)}"));
            Assert.Equal(UnstyledGlobals, globals);
        }

        [Fact]
        public void XlsStylesRecordsInOrder()
        {
            byte[] file = WriteStyledCells(new CellStyle { Bold = true, FillColor = Color(1), Left = new CellBorder(ExcelBorderStyle.Thin) });
            List<int> ids = Globals(file).Select(r => r.Id).ToList();
            int lastFont = ids.LastIndexOf(StyleProbe.BiffFont);
            int firstXf = ids.IndexOf(StyleProbe.BiffXf);
            int lastXf = ids.LastIndexOf(StyleProbe.BiffXf);
            int style = ids.IndexOf(StyleProbe.BiffStyle);
            int palette = ids.IndexOf(StyleProbe.BiffPalette);
            int boundSheet = ids.IndexOf(BoundSheet);
            Assert.True(lastFont < firstXf && lastXf < style && style < palette && palette < boundSheet);
            Assert.Equal(5, ids.Count(id => id == StyleProbe.BiffFont));
        }

        [Fact]
        public void BoldFromBeforeThisChangeIsNowWritten()
        {
            byte[] file = WriteStyledCells(new CellStyle { Bold = true });
            Assert.Equal(Defaults with { Bold = true }, StyleProbe.Xls(file, 0, 0));
        }

        [Fact]
        public void FontIndexFourIsNeverReferenced()
        {
            byte[] file = WriteStyledCells(new CellStyle { Bold = true }, new CellStyle { Italic = true }, new CellStyle { FontSize = 20 });
            List<XlsRecord> records = StyleProbe.XlsRecords(file);
            Assert.Equal(4 + 3, StyleProbe.XlsPayloads(records, StyleProbe.BiffFont).Length);
            HashSet<int> fontIndexes = StyleProbe.XlsPayloads(records, StyleProbe.BiffXf).Select(xf => (int)BitConverter.ToUInt16(xf, 0)).ToHashSet();
            Assert.Equal([0, 5, 6, 7], fontIndexes.Order());
            Assert.Equal(Defaults with { Italic = true }, StyleProbe.Xls(file, 1, 0));
            Assert.Equal(Defaults with { FontSize = 20 }, StyleProbe.Xls(file, 2, 0));
        }

        [Fact]
        public void FiftySeventhColorThrowsFromAddStyle()
        {
            using var wb = XlsWorkbookWriter.Create(new MemoryStream());
            for (int i = 0; i < 56; i++)
            {
                wb.AddStyle(new CellStyle { FillColor = Color(i) });
            }
            var ex = Assert.Throws<ExcelLimitExceededException>(() => wb.AddStyle(new CellStyle { FillColor = Color(56) }));
            Assert.Equal("Colors", ex.LimitName);
            Assert.Equal(56L, ex.Limit);
        }

        [Fact]
        public void PaletteOmittedWhenNoColorIsUsed()
        {
            byte[] file = WriteStyledCells(new CellStyle { Bold = true });
            Assert.DoesNotContain(StyleProbe.XlsRecords(file), r => r.Id == StyleProbe.BiffPalette);
        }
    }
}
