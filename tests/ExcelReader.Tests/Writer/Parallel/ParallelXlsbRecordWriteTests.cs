using System.Globalization;
using System.IO.Compression;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Tests.Writer.Zip;
using Item = ExcelReader.Tests.Writer.Parallel.ParallelRecordWriteTests.Item;

namespace ExcelReader.Tests.Writer.Parallel
{
    public class ParallelXlsbRecordWriteTests
    {
        public static TheoryData<CompressionLevel, bool, bool> Modes()
        {
            return ParallelRecordWriteTests.Modes();
        }

        private static async Task<byte[]> WriteAsync(XlsbWriterOptions options, Func<XlsbSheetWriter, Task> body, bool seekable = true)
        {
            var ms = new MemoryStream();
            Stream target = seekable ? ms : new ForwardOnlyWriteStream(ms);
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(target, leaveOpen: true, options))
            {
                XlsbSheetWriter sheet = wb.AddSheet("Data");
                await body(sheet);
                await sheet.EndAsync(TestContext.Current.CancellationToken);
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }
            return ms.ToArray();
        }

        private static ParallelWriteOptions Parallel(int degree, int batchRows = 500)
        {
            return new ParallelWriteOptions { DegreeOfParallelism = degree, BatchRows = batchRows };
        }

        private static List<string[]> Read(byte[] package)
        {
            using XlsbWorkbook workbook = Excel.FromXlsb(new MemoryStream(package));
            return ParallelRecordWriteTests.Cells(workbook);
        }

        private static void Fill(XlsbRowWriter row, Item item)
        {
            ParallelRecordWriteTests.Fill(row, item);
        }

        [Theory]
        [MemberData(nameof(Modes))]
        public async Task ParallelGivesTheSameCellsAsSerial(CompressionLevel level, bool prefetch, bool sharedStrings)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = ParallelRecordWriteTests.Items(6_300);
            var options = new XlsbWriterOptions { Compression = level, PrefetchWrite = prefetch, UseSharedStrings = sharedStrings };
            byte[] serial = await WriteAsync(options, async sheet => await sheet.WriteRecordsAsync(items, Fill, ct));
            foreach (bool seekable in new[] { true, false })
            {
                byte[] parallel = await WriteAsync(options, async sheet =>
                    await sheet.WriteRecordsParallelAsync(items, Fill, Parallel(4), ct), seekable);

                Assert.Equal(Read(serial), Read(parallel));
                ParallelRecordWriteTests.Entry(parallel, "xl/worksheets/sheet1.bin");
            }
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(500)]
        [InlineData(501)]
        [InlineData(1_733)]
        public async Task BatchEdges(int count)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = ParallelRecordWriteTests.Items(count);
            byte[] package = await WriteAsync(XlsbWriterOptions.Default, async sheet =>
                await sheet.WriteRecordsParallelAsync(items, Fill, Parallel(3), ct));

            List<string[]> rows = Read(package);
            Assert.Equal(count, rows.Count);
            for (int i = 0; i < count; i++)
            {
                Assert.Equal(i.ToString(CultureInfo.InvariantCulture), rows[i][0]);
                Assert.Equal(items[i].Name, rows[i][1]);
            }
        }

        [Fact]
        public async Task ParallelAsTheFirstWriteOnTheSheet()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = ParallelRecordWriteTests.Items(1_200);
            byte[] package = await WriteAsync(XlsbWriterOptions.Default with { PrefetchWrite = true }, async sheet =>
                await sheet.WriteRecordsParallelAsync(items, Fill, Parallel(4), ct));

            Assert.Equal(1_200, Read(package).Count);
        }

        [Theory]
        [InlineData(CompressionLevel.Fastest, false)]
        [InlineData(CompressionLevel.Optimal, false)]
        [InlineData(CompressionLevel.NoCompression, false)]
        [InlineData(CompressionLevel.Fastest, true)]
        [InlineData(CompressionLevel.Optimal, true)]
        [InlineData(CompressionLevel.NoCompression, true)]
        public async Task SerialRowsBeforeAndAfterSurvive(CompressionLevel level, bool prefetch)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = ParallelRecordWriteTests.Items(1_300);
            byte[] package = await WriteAsync(new XlsbWriterOptions { Compression = level, PrefetchWrite = prefetch }, async sheet =>
            {
                await sheet.WriteRecordsAsync(items, Fill, ct);
                await sheet.WriteRecordsParallelAsync(items, Fill, Parallel(4), ct);
                await sheet.WriteRecordsAsync(items, Fill, ct);
                await sheet.WriteRecordsParallelAsync(items, Fill, Parallel(2), ct);
                using XlsbRowWriter last = sheet.StartRow();
                last.Write("end");
            });

            List<string[]> rows = Read(package);
            Assert.Equal((4 * 1_300) + 1, rows.Count);
            for (int block = 0; block < 4; block++)
            {
                Assert.Equal("0", rows[block * 1_300][0]);
                Assert.Equal("1299", rows[(block * 1_300) + 1_299][0]);
            }
            Assert.Equal("end", rows[^1][0]);
        }

        [Fact]
        public async Task ALayoutAndATableMatchSerial()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = ParallelRecordWriteTests.Items(900);
            byte[] serial = await WriteAsync(XlsbWriterOptions.Default, async sheet =>
                await sheet.WriteTableAsync(items, "Records", ExcelRecordLayout.FromAttributes<Item>(), ct: ct));
            byte[] parallel = await WriteAsync(XlsbWriterOptions.Default, async sheet =>
                await sheet.WriteTableParallelAsync(items, "Records", ExcelRecordLayout.FromAttributes<Item>(), options: Parallel(4, 100), ct: ct));

            Assert.Equal(Read(serial), Read(parallel));
            Assert.Equal(ParallelRecordWriteTests.Entry(serial, "xl/tables/table1.bin"), ParallelRecordWriteTests.Entry(parallel, "xl/tables/table1.bin"));
        }

        [Fact]
        public async Task SharedStringsGiveEveryCellItsOwnText()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = ParallelRecordWriteTests.Items(4_000, distinctNames: 613);
            for (int round = 0; round < 25; round++)
            {
                byte[] package = await WriteAsync(new XlsbWriterOptions { UseSharedStrings = true }, async sheet =>
                    await sheet.WriteRecordsParallelAsync(items, Fill, Parallel(8, 50), ct));

                List<string[]> rows = Read(package);
                Assert.Equal(items.Count, rows.Count);
                for (int i = 0; i < items.Count; i++)
                {
                    Assert.Equal(items[i].Name, rows[i][1]);
                }
                Assert.Equal((4_000u, 613u), SharedStringCounts(package));
            }
        }

        // BrtBeginSst opens the part: a two-byte record id, a one-byte size, then the reference count
        // and the unique count as 32-bit integers.
        private static (uint References, uint Unique) SharedStringCounts(byte[] package)
        {
            byte[] part = ParallelRecordWriteTests.Entry(package, "xl/sharedStrings.bin");
            return (BitConverter.ToUInt32(part, 3), BitConverter.ToUInt32(part, 7));
        }

        [Fact]
        public async Task SharedStringCountsMatchASerialWrite()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = ParallelRecordWriteTests.Items(4_000, distinctNames: 613);
            var options = new XlsbWriterOptions { UseSharedStrings = true };
            byte[] serial = await WriteAsync(options, async sheet => await sheet.WriteRecordsAsync(items, Fill, ct));
            byte[] parallel = await WriteAsync(options, async sheet =>
                await sheet.WriteRecordsParallelAsync(items, Fill, Parallel(8, 50), ct));

            Assert.Equal(SharedStringCounts(serial), SharedStringCounts(parallel));
        }

        [Fact]
        public async Task AThrowingRowFaultsTheSheet()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = ParallelRecordWriteTests.Items(5_000);
            static void Throwing(XlsbRowWriter row, Item item)
            {
                row.Write(item.Id);
                if (item.Id == 2_345)
                {
                    throw new FormatException("row 2345");
                }
            }
            var ms = new MemoryStream();
            XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true);
            XlsbSheetWriter sheet = wb.AddSheet("Data");

            await Assert.ThrowsAsync<FormatException>(
                async () => await sheet.WriteRecordsParallelAsync(items, Throwing, Parallel(4, 100), ct));

            Assert.True(sheet.ResourcesReleased);
            Assert.Throws<ObjectDisposedException>(() => sheet.StartRow());
            await wb.DisposeAsync();
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task AFailingOutputSurfacesItsOwnException(bool prefetch)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = ParallelRecordWriteTests.Items(20_000);
            using var target = new FailingWriteStream(limit: 4096);
            XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(target, leaveOpen: true, new XlsbWriterOptions { PrefetchWrite = prefetch });
            XlsbSheetWriter sheet = wb.AddSheet("Data");

            IOException failure = await Assert.ThrowsAsync<IOException>(
                async () => await sheet.WriteRecordsParallelAsync(items, Fill, Parallel(4), ct));

            Assert.Equal("disk full", failure.Message);
            Assert.True(sheet.ResourcesReleased);
            await wb.DisposeAsync();
        }
    }
}
