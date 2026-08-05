using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Tests
{
    public class WriterStyleTests
    {
        private static async Task<MemoryStream> WriteXlsxAsync(Func<XlsxWorkbookWriter, Task> build)
        {
            var ms = new MemoryStream();
            await using var wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            await build(wb);
            await wb.EndAsync(TestContext.Current.CancellationToken);
            ms.Position = 0;
            return ms;
        }

        private static string ReadZipEntryText(MemoryStream xlsx, string entryName)
        {
            xlsx.Position = 0;
            using var zip = new ZipArchive(xlsx, ZipArchiveMode.Read, leaveOpen: true);
            using StreamReader reader = new(zip.GetEntry(entryName)!.Open());
            return reader.ReadToEnd();
        }

        [Fact]
        public async Task AddStyleReturnsZeroForDefaultStyle()
        {
            var ms = new MemoryStream();
            await using var wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken);
            Assert.Equal(0, wb.AddStyle(default));
        }

        [Fact]
        public async Task AddStyleDeduplicatesEquivalentStyles()
        {
            var ms = new MemoryStream();
            await using var wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken);
            var style = new CellStyle { NumberFormat = "0.00", Bold = true };
            int first = wb.AddStyle(style);
            int second = wb.AddStyle(new CellStyle { NumberFormat = "0.00", Bold = true });
            int different = wb.AddStyle(new CellStyle { NumberFormat = "0.00%" });
            Assert.Equal(first, second);
            Assert.NotEqual(first, different);
        }

        [Fact]
        public async Task CustomNumberFormatIdStartsAt164()
        {
            await using MemoryStream ms = await WriteXlsxAsync(async wb =>
            {
                wb.AddStyle(new CellStyle { NumberFormat = "R$ #,##0.00" });
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });
            string styles = ReadZipEntryText(ms, "xl/styles.xml");
            Assert.Contains("numFmtId=\"164\" formatCode=\"R$ #,##0.00\"", styles, StringComparison.Ordinal);
        }

        [Fact]
        public async Task DateCellsKeepStyleIndexOneXlsx()
        {
            var date = new DateTime(2024, 5, 6, 0, 0, 0, DateTimeKind.Unspecified);
            await using MemoryStream ms = await WriteXlsxAsync(async wb =>
            {
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write(date);
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(1, e.Current[0].StyleIndex);
        }

        [Fact]
        public async Task DateCellsKeepStyleIndexOneXlsb()
        {
            var date = new DateTime(2024, 5, 6, 0, 0, 0, DateTimeKind.Unspecified);
            var ms = new MemoryStream();
            await using (var wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken))
            {
                await wb.StartAsync(TestContext.Current.CancellationToken);
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write(date);
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }
            ms.Position = 0;
            await using var reader = Excel.FromXlsb(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(1, e.Current[0].StyleIndex);
        }

        [Fact]
        public async Task DateCellsKeepStyleIndexOneXls()
        {
            var date = new DateTime(2024, 5, 6, 0, 0, 0, DateTimeKind.Unspecified);
            var ms = new MemoryStream();
            await using (var wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                wb.Start();
                XlsSheetWriter sheet = wb.AddSheet("Sheet1");
                sheet.Start();
                using (var row = sheet.StartRow())
                {
                    row.Write(date);
                }
                sheet.End();
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }
            ms.Position = 0;
            using var reader = Excel.FromXls(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            // XlsReader reports the file's raw XF index (unlike Xlsx/Xlsb, XLS offsets every style by
            // the 16 builtin style XFs written ahead of the general/date pair), not the abstract
            // AddStyle index — DateXf is that raw index for the builtin date style.
            Assert.Equal(XlsGlobals.DateXf, e.Current[0].StyleIndex);
        }

        [Fact]
        public async Task ColumnStyleAfterSheetStartThrowsXlsx()
        {
            await using var ms = new MemoryStream();
            await using var wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.StartAsync(TestContext.Current.CancellationToken);
            Assert.Throws<InvalidOperationException>(() => sheet.SetColumnStyle(0, 1));
            Assert.Throws<InvalidOperationException>(() => sheet.SetColumnWidth(0, 12));
        }

        [Fact]
        public async Task ColumnStyleAfterSheetStartThrowsXlsb()
        {
            await using var ms = new MemoryStream();
            await using var wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.StartAsync(TestContext.Current.CancellationToken);
            Assert.Throws<InvalidOperationException>(() => sheet.SetColumnStyle(0, 1));
            Assert.Throws<InvalidOperationException>(() => sheet.SetColumnWidth(0, 12));
        }

        [Fact]
        public async Task ColumnStyleAfterSheetStartThrowsXls()
        {
            await using var ms = new MemoryStream();
            await using var wb = XlsWorkbookWriter.Create(ms, leaveOpen: true);
            wb.Start();
            XlsSheetWriter sheet = wb.AddSheet("Sheet1");
            sheet.Start();
            Assert.Throws<InvalidOperationException>(() => sheet.SetColumnStyle(0, 1));
            Assert.Throws<InvalidOperationException>(() => sheet.SetColumnWidth(0, 12));
        }

        [Fact]
        public async Task ColumnStyleAfterSheetStartIsNoOpForCsv()
        {
            await using var ms = new MemoryStream();
            CsvWorkbookWriter wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            CsvSheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.StartAsync(TestContext.Current.CancellationToken);
            sheet.SetColumnStyle(0, 1);
            sheet.SetColumnWidth(0, 12);
            await using (CsvRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
            {
                row.Write(42);
            }
            await wb.EndAsync(TestContext.Current.CancellationToken);
            Assert.Equal("42\r\n", Encoding.UTF8.GetString(ms.ToArray()));
        }

        [Fact]
        public async Task ColumnStyleRoundTripsThroughReaderXlsx()
        {
            await using MemoryStream ms = await WriteXlsxAsync(async wb =>
            {
                int styleId = wb.AddStyle(new CellStyle { NumberFormat = "0.00" });
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
                sheet.SetColumnStyle(0, styleId);
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write(42);
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(2, e.Current[0].StyleIndex);
        }

        [Fact]
        public async Task ColumnStyleRoundTripsThroughReaderXlsb()
        {
            var ms = new MemoryStream();
            int styleId;
            await using (var wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken))
            {
                styleId = wb.AddStyle(new CellStyle { NumberFormat = "0.00" });
                await wb.StartAsync(TestContext.Current.CancellationToken);
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                sheet.SetColumnStyle(0, styleId);
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write(42);
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }
            ms.Position = 0;
            await using var reader = Excel.FromXlsb(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(styleId, e.Current[0].StyleIndex);
        }

        [Fact]
        public async Task ColumnStyleRoundTripsThroughReaderXls()
        {
            var ms = new MemoryStream();
            int styleId;
            await using (var wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                styleId = wb.AddStyle(new CellStyle { NumberFormat = "0.00" });
                wb.Start();
                XlsSheetWriter sheet = wb.AddSheet("Sheet1");
                sheet.SetColumnStyle(0, styleId);
                sheet.Start();
                using (var row = sheet.StartRow())
                {
                    row.Write(42);
                }
                sheet.End();
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }
            ms.Position = 0;
            using var reader = Excel.FromXls(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(XlsGlobals.CustomXf(styleId), e.Current[0].StyleIndex);
        }

        [Fact]
        public async Task ColumnStyleIsNoOpForCsv()
        {
            await using var ms = new MemoryStream();
            CsvWorkbookWriter wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);
            int styleId = wb.AddStyle(new CellStyle { NumberFormat = "0.00" });
            Assert.Equal(0, styleId);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            CsvSheetWriter sheet = wb.AddSheet("Sheet1");
            sheet.SetColumnStyle(0, 1);
            await using (CsvRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
            {
                row.Write(42);
            }
            await wb.EndAsync(TestContext.Current.CancellationToken);
            Assert.Equal("42\r\n", Encoding.UTF8.GetString(ms.ToArray()));
        }

        [Fact]
        public async Task RowStyleRoundTripsThroughReaderXlsx()
        {
            await using MemoryStream ms = await WriteXlsxAsync(async wb =>
            {
                int styleId = wb.AddStyle(new CellStyle { Bold = true });
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsxRowWriter row = await sheet.StartRowAsync(styleId, TestContext.Current.CancellationToken))
                {
                    row.Write("Header");
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(2, e.Current[0].StyleIndex);
        }

        [Fact]
        public async Task RowStyleRoundTripsThroughReaderXlsb()
        {
            var ms = new MemoryStream();
            int styleId;
            await using (var wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken))
            {
                styleId = wb.AddStyle(new CellStyle { Bold = true });
                await wb.StartAsync(TestContext.Current.CancellationToken);
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(styleId, TestContext.Current.CancellationToken))
                {
                    row.Write("Header");
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }
            ms.Position = 0;
            await using var reader = Excel.FromXlsb(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(styleId, e.Current[0].StyleIndex);
        }

        [Fact]
        public async Task RowStyleRoundTripsThroughReaderXls()
        {
            var ms = new MemoryStream();
            int styleId;
            await using (var wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                styleId = wb.AddStyle(new CellStyle { Bold = true });
                wb.Start();
                XlsSheetWriter sheet = wb.AddSheet("Sheet1");
                sheet.Start();
                using (var row = sheet.StartRow(styleId))
                {
                    row.Write("Header");
                }
                sheet.End();
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }
            ms.Position = 0;
            using var reader = Excel.FromXls(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(XlsGlobals.CustomXf(styleId), e.Current[0].StyleIndex);
        }

        [Fact]
        public async Task RowStyleIsNoOpForCsv()
        {
            await using var ms = new MemoryStream();
            CsvWorkbookWriter wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);
            int styleId = wb.AddStyle(new CellStyle { Bold = true });
            await wb.StartAsync(TestContext.Current.CancellationToken);
            CsvSheetWriter sheet = wb.AddSheet("Sheet1");
            await using (CsvRowWriter row = await sheet.StartRowAsync(styleId, TestContext.Current.CancellationToken))
            {
                row.Write("Header");
            }
            await wb.EndAsync(TestContext.Current.CancellationToken);
            Assert.Equal("Header\r\n", Encoding.UTF8.GetString(ms.ToArray()));
        }

        [Fact]
        public async Task StyledWorkbookOpensWithoutRepairXlsx()
        {
            await using MemoryStream ms = await WriteXlsxAsync(async wb =>
            {
                wb.AddStyle(new CellStyle { NumberFormat = "R$ #,##0.00" });
                XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });
            string styles = ReadZipEntryText(ms, "xl/styles.xml");
            const string expected =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                "<numFmts count=\"2\"><numFmt numFmtId=\"14\" formatCode=\"mm-dd-yy\"/><numFmt numFmtId=\"164\" formatCode=\"R$ #,##0.00\"/></numFmts>" +
                "<fonts count=\"1\"><font/></fonts>" +
                "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>" +
                "<borders count=\"1\"><border/></borders>" +
                "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
                "<cellXfs count=\"3\">" +
                "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
                "<xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
                "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
                "</cellXfs>" +
                "</styleSheet>";
            Assert.Equal(expected, styles);
        }

        [Fact]
        public async Task InvalidStyleArgumentsThrowXlsx()
        {
            await using var ms = new MemoryStream();
            await using var wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            XlsxSheetWriter sheet = wb.AddSheet("Sheet1");

            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnStyle(-1, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnStyle(0, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnStyle(0, 1000));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidth(-1, 12));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidth(0, -1));

            await sheet.StartAsync(TestContext.Current.CancellationToken);
            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(async () => await sheet.StartRowAsync(-1, TestContext.Current.CancellationToken));
            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(async () => await sheet.StartRowAsync(1000, TestContext.Current.CancellationToken));
        }

        [Fact]
        public async Task InvalidStyleArgumentsThrowXlsb()
        {
            await using var ms = new MemoryStream();
            await using var wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            XlsbSheetWriter sheet = wb.AddSheet("Sheet1");

            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnStyle(-1, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnStyle(0, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnStyle(0, 1000));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidth(-1, 12));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidth(0, -1));

            await sheet.StartAsync(TestContext.Current.CancellationToken);
            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(async () => await sheet.StartRowAsync(-1, TestContext.Current.CancellationToken));
            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(async () => await sheet.StartRowAsync(1000, TestContext.Current.CancellationToken));
        }

        [Fact]
        public async Task InvalidStyleArgumentsThrowXls()
        {
            await using var ms = new MemoryStream();
            await using var wb = XlsWorkbookWriter.Create(ms, leaveOpen: true);
            wb.Start();
            XlsSheetWriter sheet = wb.AddSheet("Sheet1");

            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnStyle(-1, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnStyle(0, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnStyle(0, 1000));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidth(-1, 12));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidth(0, -1));

            sheet.Start();
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.StartRow(-1));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.StartRow(1000));
        }

        // CSV only rejects negative arguments — a styleId that no format's AddStyle ever handed out
        // still no-ops here rather than throwing (CsvWorkbookWriter.AddStyle's doc comment / the
        // SetColumnStyle no-op above explain why: CSV never writes styleId anywhere, so a caller
        // sharing one styleId literal across all four formats must not have the CSV leg alone reject it).
        [Fact]
        public async Task InvalidStyleArgumentsThrowCsv()
        {
            await using var ms = new MemoryStream();
            CsvWorkbookWriter wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            CsvSheetWriter sheet = wb.AddSheet("Sheet1");

            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnStyle(-1, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnStyle(0, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidth(-1, 12));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidth(0, -1));

            await sheet.StartAsync(TestContext.Current.CancellationToken);
            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(async () => await sheet.StartRowAsync(-1, TestContext.Current.CancellationToken));
            // Unregistered but non-negative: no-op, not an exception.
            await using (CsvRowWriter row = await sheet.StartRowAsync(1000, TestContext.Current.CancellationToken))
            {
                row.Write(1);
            }
        }
    }
}
