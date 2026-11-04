using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using ExcelReader.NetFramework.Reader;
using ExcelReader.NetFramework.Writer;

namespace ExcelReader.NetFramework.Tests.Writer
{
    public sealed class StyledRow
    {
        public string? Name { get; set; }

        [ExcelStyle(Bold = true)]
        public DateTime When { get; set; }

        [ExcelStyle(NumberFormat = "#,##0.000", FillColor = 0xFFFF00)]
        public double Price { get; set; }
    }

    public sealed class StyleTests
    {
        private static string Entry(byte[] xlsx, string name)
        {
            using ZipArchive zip = new ZipArchive(new MemoryStream(xlsx), ZipArchiveMode.Read);
            using StreamReader reader = new StreamReader(zip.GetEntry(name)!.Open());
            return reader.ReadToEnd();
        }

        private static int[][] CellStyles(byte[] xlsx)
        {
            string sheet = Entry(xlsx, "xl/worksheets/sheet1.xml");
            TimeSpan timeout = TimeSpan.FromSeconds(5);
            List<int[]> rows = new List<int[]>();
            foreach (Match row in Regex.Matches(sheet, "<row[^>]*>(?<cells>.*?)</row>", RegexOptions.Singleline | RegexOptions.ExplicitCapture, timeout))
            {
                List<int> cells = new List<int>();
                foreach (Match cell in Regex.Matches(row.Groups["cells"].Value, "<c(?<attrs>[^>]*)>", RegexOptions.ExplicitCapture, timeout))
                {
                    Match style = Regex.Match(cell.Groups["attrs"].Value, "s=\"(?<id>[0-9]+)\"", RegexOptions.ExplicitCapture, timeout);
                    cells.Add(style.Success ? int.Parse(style.Groups["id"].Value, CultureInfo.InvariantCulture) : 0);
                }

                rows.Add(cells.ToArray());
            }

            return rows.ToArray();
        }

        private static int Of(int[][] cells, string reference)
        {
            return cells[reference[1] - '1'][reference[0] - 'A'];
        }

        private static List<StyledRow> Rows()
        {
            return new List<StyledRow>
            {
                new StyledRow { Name = "a", When = new DateTime(2026, 3, 15, 10, 20, 30), Price = 1.5 },
                new StyledRow { Name = "b", When = new DateTime(2026, 3, 16, 1, 2, 3), Price = 2.5 },
            };
        }

        [Fact]
        public void AddStyle_DedupesAndReachesStylesXml()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            CellStyle style = new CellStyle { Bold = true, NumberFormat = "0.0%", FontName = "Consolas", FontSize = 14, FillColor = 0x00FF00, Left = new CellBorder(ExcelBorderStyle.Thin) };
            int first = writer.AddStyle(style);
            int second = writer.AddStyle(new CellStyle { Bold = true, NumberFormat = "0.0%", FontName = "Consolas", FontSize = 14, FillColor = 0x00FF00, Left = new CellBorder(ExcelBorderStyle.Thin) });
            writer.StartSheet("S");
            writer.StartRow();
            writer.StyleNextCell(first);
            writer.Write(0.25);
            writer.EndRow();
            writer.EndSheet();
            byte[] bytes = writer.CloseToArray();

            string styles = Entry(bytes, "xl/styles.xml");
            Assert.True(first > 0);
            Assert.Equal(first, second);
            Assert.Contains("0.0%", styles, StringComparison.Ordinal);
            Assert.Contains("Consolas", styles, StringComparison.Ordinal);
            Assert.Contains("00FF00", styles, StringComparison.Ordinal);
            Assert.Equal(first, Of(CellStyles(bytes), "A1"));
        }

        [Fact]
        public void DefaultStyle_IsGeneralAndColorsDefaultToUnset()
        {
            CellStyle style = new CellStyle();
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);

            Assert.Equal(-1, style.FontColor);
            Assert.Equal(-1, style.FillColor);
            Assert.Equal(0, writer.AddStyle(style));
        }

        [Fact]
        public void ColumnRowAndNextCellStyles_ApplyByPrecedence()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            int column = writer.AddStyle(new CellStyle { Italic = true });
            int row = writer.AddStyle(new CellStyle { Bold = true });
            int cell = writer.AddStyle(new CellStyle { Underline = true });
            writer.StartSheet("S");
            writer.SetColumnStyle(0, column);
            writer.SetColumnWidth(0, 30);
            writer.StartRow();
            writer.Write("col");
            writer.EndRow();
            writer.StartRow(row, 40);
            writer.Write("row");
            writer.EndRow();
            writer.StartRow();
            writer.StyleNextCell(cell);
            writer.Write("cell");
            writer.Write("plain");
            writer.EndRow();
            writer.EndSheet();
            byte[] bytes = writer.CloseToArray();

            int[][] cells = CellStyles(bytes);
            string sheet = Entry(bytes, "xl/worksheets/sheet1.xml");
            Assert.Equal(column, Of(cells, "A1"));
            Assert.Equal(row, Of(cells, "A2"));
            Assert.Equal(cell, Of(cells, "A3"));
            Assert.Equal(0, Of(cells, "B3"));
            Assert.Contains("width=\"30", sheet, StringComparison.Ordinal);
            Assert.Contains("ht=\"40\"", sheet, StringComparison.Ordinal);
        }

        [Fact]
        public void StyledNextCell_WithEmpty_WritesStyledBlank()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            int id = writer.AddStyle(new CellStyle { FillColor = 0xFF0000 });
            writer.StartSheet("S");
            writer.StartRow();
            writer.StyleNextCell(id);
            writer.WriteEmpty();
            writer.EndRow();
            writer.EndSheet();

            Assert.Equal(id, Of(CellStyles(writer.CloseToArray()), "A1"));
        }

        [Fact]
        public void WriteSheet_StyleAttributeKeepsDateAndTime()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            writer.WriteSheet("S", Rows());
            byte[] bytes = writer.CloseToArray();

            using ExcelWorkbook workbook = ExcelWorkbook.Open(bytes, new ExcelOpenOptions { Format = ExcelFormat.Xlsx });
            StyledRow[] read = workbook.Sheet(0).Read<StyledRow>();
            int[][] cells = CellStyles(bytes);
            Assert.Equal(Rows()[0].When, read[0].When);
            Assert.Equal(1.5, read[0].Price);
            Assert.Contains("#,##0.000", Entry(bytes, "xl/styles.xml"), StringComparison.Ordinal);
            Assert.Contains("FFFF00", Entry(bytes, "xl/styles.xml"), StringComparison.Ordinal);
            Assert.Equal(0, Of(cells, "A1"));
            Assert.NotEqual(Of(cells, "B2"), Of(cells, "B1"));
            Assert.Equal(0, Of(cells, "C1"));
            Assert.NotEqual(0, Of(cells, "B2"));
            Assert.NotEqual(0, Of(cells, "C2"));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void WriteSheet_HeaderStyleStylesOnlyHeader(bool asTable)
        {
            ExcelTableOptions? table = asTable ? new ExcelTableOptions { Name = "T" } : null;
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            writer.WriteSheet("S", Rows(), table, new CellStyle { Bold = true, FillColor = 0xCCCCCC });
            byte[] bytes = writer.CloseToArray();

            int[][] cells = CellStyles(bytes);
            Assert.NotEqual(0, Of(cells, "A1"));
            Assert.NotEqual(0, Of(cells, "B1"));
            Assert.NotEqual(0, Of(cells, "C1"));
            Assert.Equal(0, Of(cells, "A2"));
            Assert.Equal(asTable, Entry(bytes, "[Content_Types].xml").Contains("table"));
            Assert.Contains("CCCCCC", Entry(bytes, "xl/styles.xml"), StringComparison.Ordinal);
        }

        [Fact]
        public void WriteSheet_HeaderStyleOnPlainRowsWithoutAttributes()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            writer.WriteSheet("S", new[] { new WriteRow { Name = "x", When = new DateTime(2026, 1, 1) } }, null, new CellStyle { Italic = true });

            Assert.Contains("<i", Entry(writer.CloseToArray(), "xl/styles.xml"), StringComparison.Ordinal);
        }

        [Fact]
        public void InvalidColor_ThrowsInvalidArgumentAndHandleStaysUsable()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);

            ExcelReaderException error = Assert.Throws<ExcelReaderException>(() => writer.AddStyle(new CellStyle { FontColor = 0x1000000 }));
            Assert.Equal(ExcelStatus.InvalidArgument, error.Status);
            Assert.Throws<ExcelReaderException>(() => writer.AddStyle(new CellStyle { Top = new CellBorder((ExcelBorderStyle)99) }));
            Assert.Throws<ExcelReaderException>(() => writer.AddStyle(new CellStyle { FillColor = -2 }));
            Assert.True(writer.AddStyle(new CellStyle { Bold = true }) > 0);
            writer.WriteSheet("S", Rows());
            Assert.NotEmpty(writer.CloseToArray());
        }

        [Fact]
        public void InvalidAttributeColor_ThrowsInvalidArgument()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);

            ExcelReaderException error = Assert.Throws<ExcelReaderException>(() => writer.WriteSheet("S", new[] { new BadStyleRow() }));
            Assert.Equal(ExcelStatus.InvalidArgument, error.Status);
        }

        [Fact]
        public void Csv_AcceptsStylesAsNoOps()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Csv);
            int id = writer.AddStyle(new CellStyle { Bold = true });
            writer.StartSheet("S");
            writer.SetColumnStyle(0, id);
            writer.StartRow(id);
            writer.StyleNextCell(id);
            writer.Write("x");
            writer.EndRow();
            writer.EndSheet();

            Assert.Equal("x", System.Text.Encoding.UTF8.GetString(writer.CloseToArray()).Trim());
        }

        private sealed class BadStyleRow
        {
            [ExcelStyle(FontColor = 0x1000000)]
            public int Value { get; set; }
        }
    }
}
