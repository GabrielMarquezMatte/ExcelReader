using System.Globalization;
using System.IO.Compression;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsx;
using ExcelReader.Tests.Writer.Zip;

namespace ExcelReader.Tests.Writer.Parallel
{
    public class ParallelRecordWriteTests
    {
        internal sealed record Item(int Id, string Name, double Value, DateTime Day);

        internal const int Columns = 4;

        internal static List<Item> Items(int count, int distinctNames = 37)
        {
            return [.. Enumerable.Range(0, count).Select(i => new Item(
                i, string.Create(CultureInfo.InvariantCulture, $"name {i % distinctNames}"), i * 1.5, new DateTime(2024, 1, 1).AddDays(i % 400)))];
        }

        internal static void Fill<TRow>(TRow row, Item item)
            where TRow : IRowWriter
        {
            row.Write(item.Id);
            row.Write(item.Name);
            row.Write(item.Value);
            row.Write(item.Day);
        }

        internal static List<string[]> Cells(IExcelWorkbook workbook, int sheet = 0)
        {
            var rows = new List<string[]>();
            using IExcelRowEnumerator enumerator = workbook.SheetAt(sheet).GetEnumerator();
            while (enumerator.MoveNext())
            {
                string[] cells = new string[Columns];
                for (int c = 0; c < Columns; c++)
                {
                    cells[c] = enumerator.Current[c].GetString();
                }
                rows.Add(cells);
            }
            return rows;
        }

        internal static byte[] Entry(byte[] package, string name)
        {
            using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            ZipArchiveEntry entry = zip.GetEntry(name)!;
            using Stream data = entry.Open();
            using var copy = new MemoryStream();
            data.CopyTo(copy);
            Assert.Equal(Crc32Tests.BitwiseCrc(copy.ToArray()), entry.Crc32);
            return copy.ToArray();
        }

        public static TheoryData<CompressionLevel, bool, bool> Modes()
        {
            var data = new TheoryData<CompressionLevel, bool, bool>();
            foreach (CompressionLevel level in Enum.GetValues<CompressionLevel>())
            {
                foreach (bool prefetch in new[] { false, true })
                {
                    data.Add(level, prefetch, false);
                    data.Add(level, prefetch, true);
                }
            }
            return data;
        }

        private static async Task<byte[]> WriteAsync(XlsxWriterOptions options, Func<XlsxSheetWriter, Task> body, bool seekable = true)
        {
            var ms = new MemoryStream();
            Stream target = seekable ? ms : new ForwardOnlyWriteStream(ms);
            await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(target, leaveOpen: true, options))
            {
                XlsxSheetWriter sheet = wb.AddSheet("Data");
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
            using XlsxWorkbook workbook = Excel.FromXlsx(new MemoryStream(package));
            return Cells(workbook);
        }

        [Theory]
        [MemberData(nameof(Modes))]
        public async Task ParallelGivesTheSameCellsAsSerial(CompressionLevel level, bool prefetch, bool sharedStrings)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = Items(6_300);
            var options = new XlsxWriterOptions { Compression = level, PrefetchWrite = prefetch, UseSharedStrings = sharedStrings };
            byte[] serial = await WriteAsync(options, async sheet =>
                await sheet.WriteRecordsAsync(items, (Action<XlsxRowWriter, Item>)Fill, ct));
            foreach (bool seekable in new[] { true, false })
            {
                byte[] parallel = await WriteAsync(options, async sheet =>
                    await sheet.WriteRecordsParallelAsync(items, Fill, Parallel(4), ct), seekable);

                Assert.Equal(Read(serial), Read(parallel));
                Entry(parallel, "xl/worksheets/sheet1.xml");
            }
        }

        [Theory]
        [InlineData(CompressionLevel.Fastest)]
        [InlineData(CompressionLevel.Optimal)]
        [InlineData(CompressionLevel.NoCompression)]
        public async Task OneWorkerWritesTheSameSheetBytesAsSerial(CompressionLevel level)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = Items(2_750);
            var options = new XlsxWriterOptions { Compression = level };
            byte[] serial = await WriteAsync(options, async sheet =>
                await sheet.WriteRecordsAsync(items, (Action<XlsxRowWriter, Item>)Fill, ct));
            byte[] parallel = await WriteAsync(options, async sheet =>
                await sheet.WriteRecordsParallelAsync(items, Fill, Parallel(1), ct));

            Assert.Equal(Entry(serial, "xl/worksheets/sheet1.xml"), Entry(parallel, "xl/worksheets/sheet1.xml"));
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
            List<Item> items = Items(count);
            byte[] package = await WriteAsync(XlsxWriterOptions.Default, async sheet =>
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
            List<Item> items = Items(1_200);
            byte[] package = await WriteAsync(XlsxWriterOptions.Default with { PrefetchWrite = true }, async sheet =>
                await sheet.WriteRecordsParallelAsync(items, Fill, Parallel(4), ct));

            Assert.Equal(1_200, Read(package).Count);
        }

        [Theory]
        [InlineData(CompressionLevel.Fastest)]
        [InlineData(CompressionLevel.Optimal)]
        [InlineData(CompressionLevel.SmallestSize)]
        [InlineData(CompressionLevel.NoCompression)]
        public async Task SerialRowsBeforeAndAfterSurvive(CompressionLevel level)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = Items(1_300);
            byte[] package = await WriteAsync(new XlsxWriterOptions { Compression = level }, async sheet =>
            {
                await sheet.WriteRecordsAsync(items, (Action<XlsxRowWriter, Item>)Fill, ct);
                await sheet.WriteRecordsParallelAsync(items, Fill, Parallel(4), ct);
                await sheet.WriteRecordsAsync(items, (Action<XlsxRowWriter, Item>)Fill, ct);
                await sheet.WriteRecordsParallelAsync(items, Fill, Parallel(2), ct);
                using XlsxRowWriter last = sheet.StartRow();
                last.Write("end");
            });

            List<string[]> rows = Read(package);
            Assert.Equal((4 * 1_300) + 1, rows.Count);
            for (int block = 0; block < 4; block++)
            {
                Assert.Equal("0", rows[block * 1_300][0]);
                Assert.Equal("1299", rows[(block * 1_300) + 1_299][0]);
                Assert.Equal(items[777].Name, rows[(block * 1_300) + 777][1]);
            }
            Assert.Equal("end", rows[^1][0]);
        }

        [Fact]
        public async Task ALayoutWritesItsHeaderFirst()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = Items(900);
            byte[] serial = await WriteAsync(XlsxWriterOptions.Default, async sheet =>
                await sheet.WriteRecordsAsync(items, ExcelRecordLayout.FromAttributes<Item>(), ct));
            byte[] parallel = await WriteAsync(XlsxWriterOptions.Default, async sheet =>
                await sheet.WriteRecordsParallelAsync(items, ExcelRecordLayout.FromAttributes<Item>(), Parallel(4, 100), ct));

            Assert.Equal(Read(serial), Read(parallel));
            Assert.Equal(901, Read(parallel).Count);
        }

        [Fact]
        public async Task AnAsyncSourceWorks()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = Items(1_050);
            async IAsyncEnumerable<Item> Source()
            {
                foreach (Item item in items)
                {
                    await Task.Yield();
                    yield return item;
                }
            }
            byte[] package = await WriteAsync(XlsxWriterOptions.Default, async sheet =>
                await sheet.WriteRecordsParallelAsync(Source(), Fill, Parallel(4, 100), ct));

            Assert.Equal(1_050, Read(package).Count);
        }

        [Fact]
        public async Task AParallelTableCoversItsRows()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = Items(820);
            byte[] serial = await WriteAsync(XlsxWriterOptions.Default, async sheet =>
                await sheet.WriteTableAsync(items, "Records", ExcelRecordLayout.FromAttributes<Item>(), ct: ct));
            byte[] parallel = await WriteAsync(XlsxWriterOptions.Default, async sheet =>
                await sheet.WriteTableParallelAsync(items, "Records", ExcelRecordLayout.FromAttributes<Item>(), options: Parallel(4, 100), ct: ct));

            Assert.Equal(Read(serial), Read(parallel));
            Assert.Equal(Entry(serial, "xl/tables/table1.xml"), Entry(parallel, "xl/tables/table1.xml"));
        }

        [Fact]
        public async Task SharedStringsGiveEveryCellItsOwnText()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = Items(4_000, distinctNames: 613);
            for (int round = 0; round < 25; round++)
            {
                byte[] package = await WriteAsync(new XlsxWriterOptions { UseSharedStrings = true }, async sheet =>
                    await sheet.WriteRecordsParallelAsync(items, Fill, Parallel(8, 50), ct));

                List<string[]> rows = Read(package);
                Assert.Equal(items.Count, rows.Count);
                for (int i = 0; i < items.Count; i++)
                {
                    Assert.Equal(items[i].Name, rows[i][1]);
                }
                string sst = System.Text.Encoding.UTF8.GetString(Entry(package, "xl/sharedStrings.xml"));
                Assert.Contains("count=\"4000\" uniqueCount=\"613\"", sst, StringComparison.Ordinal);
            }
        }

        [Fact]
        public async Task AThrowingRowFaultsTheSheetAndStopsTheWorkers()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            List<Item> items = Items(5_000);
            int running = 0;
            void Throwing(XlsxRowWriter row, Item item)
            {
                Interlocked.Increment(ref running);
                try
                {
                    row.Write(item.Id);
                    if (item.Id == 2_345)
                    {
                        throw new FormatException("row 2345");
                    }
                    row.Write(item.Name);
                }
                finally
                {
                    Interlocked.Decrement(ref running);
                }
            }
            var ms = new MemoryStream();
            XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
            XlsxSheetWriter sheet = wb.AddSheet("Data");

            FormatException failure = await Assert.ThrowsAsync<FormatException>(
                async () => await sheet.WriteRecordsParallelAsync(items, Throwing, Parallel(4, 100), ct));

            Assert.Equal("row 2345", failure.Message);
            Assert.Equal(0, Volatile.Read(ref running));
            Assert.True(sheet.ResourcesReleased);
            Assert.Throws<ObjectDisposedException>(() => sheet.StartRow());
            await wb.DisposeAsync();
        }

        [Fact]
        public async Task TheSheetRefusesRowsWhileAParallelWriteRuns()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var ms = new MemoryStream();
            await using XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
            XlsxSheetWriter sheet = wb.AddSheet("Data");
            InvalidOperationException? seen = null;
            IEnumerable<Item> Source()
            {
                foreach (Item item in Items(50))
                {
                    if (item.Id == 20)
                    {
                        seen = Assert.Throws<InvalidOperationException>(() => sheet.StartRow());
                    }
                    yield return item;
                }
            }

            await sheet.WriteRecordsParallelAsync(Source(), Fill, Parallel(2, 10), ct);

            Assert.NotNull(seen);
            await sheet.EndAsync(ct);
            await wb.EndAsync(ct);
        }

        [Fact]
        public async Task InvalidArgumentsThrowBeforeAnythingIsWritten()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var ms = new MemoryStream();
            await using XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
            XlsxSheetWriter sheet = wb.AddSheet("Data");
            List<Item> items = Items(3);

            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(async () =>
                await sheet.WriteRecordsParallelAsync(items, ExcelRecordLayout.FromAttributes<Item>(), new ParallelWriteOptions { BatchRows = 0 }, ct));
            await Assert.ThrowsAsync<ArgumentNullException>(async () =>
                await sheet.WriteRecordsParallelAsync((IEnumerable<Item>)null!, Fill, ct: ct));
            await sheet.WriteRecordsParallelAsync(items, Fill, ct: ct);
            await sheet.EndAsync(ct);
            await wb.EndAsync(ct);

            Assert.Equal(3, Read(ms.ToArray()).Count);
        }
    }
}
