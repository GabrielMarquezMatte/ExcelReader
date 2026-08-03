using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    // WriteOffloadStream overlaps a ZIP entry's deflate with row-serialization on a background thread
    // (the write-side mirror of PrefetchDecompressionTests). Each test here pins one specific way that
    // can go wrong: a cell-content mismatch versus the non-offloaded path, or a sheet too small to ever
    // engage the offload path at all (XLSB only spills past SpillThreshold).
    //
    // Deliberately NOT a raw byte-for-byte comparison of the two builds' zip output: ZipArchiveEntry
    // stamps LastWriteTime from DateTime.Now at CreateEntry time, so two independent builds a few
    // milliseconds apart can legitimately cross a 2-second DOS-time rounding boundary and differ in a
    // header byte or two — a real, pre-existing quirk of System.IO.Compression, not something
    // WriteOffloadStream introduces. Comparing decoded cell content (as PrefetchDecompressionTests does
    // on the read side) is the correctness property that actually matters here.
    public class WriteOffloadTests
    {
        // Enough rows that the XLSX row buffer crosses its 64 KiB flush threshold and the XLSB record
        // buffer crosses its 64 KiB spill threshold many times over, not just once.
        private const int Rows = 20_000;

        [Fact]
        public async Task XlsxPrefetchWriteReadsBackIdenticalCells()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = await BuildXlsxAsync(prefetchWrite: true, ct);

            using MemoryStream ms = new(bytes, writable: false);
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            int rowIndex = 0;
            while (e.MoveNext())
            {
                Row row = e.Current;
                Assert.Equal(rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture), row[0].GetString());
                Assert.Equal($"row {rowIndex} text", row[1].GetString());
                rowIndex++;
            }
            Assert.Equal(Rows, rowIndex);
        }

        [Fact]
        public async Task XlsbPrefetchWriteReadsBackIdenticalCells()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = await BuildXlsbAsync(prefetchWrite: true, ct);

            using MemoryStream ms = new(bytes, writable: false);
            using XlsbReader reader = Excel.FromXlsb(ms);
            using XlsbReader.Enumerator e = reader.GetEnumerator();
            int rowIndex = 0;
            while (e.MoveNext())
            {
                Row row = e.Current;
                Assert.True(row[0].TryGetDouble(out double id));
                Assert.Equal(rowIndex, (int)id);
                Assert.Equal($"row {rowIndex} text", row[1].GetString());
                rowIndex++;
            }
            Assert.Equal(Rows, rowIndex);
        }

        // Uses the sync row-write API (StartRow/RowWriter.Dispose), which is the path that exercises
        // WriteOffloadStream's synchronous Write(ReadOnlySpan<byte>) override rather than WriteAsync.
        [Fact]
        public async Task XlsxPrefetchWriteWorksThroughSyncRowApi()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = new();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true, prefetchWrite: true, ct: ct))
            {
                await wb.StartAsync(ct);
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(ct);
                for (int r = 0; r < Rows; r++)
                {
                    using XlsxRowWriter row = sheet.StartRow(ct);
                    row.Write(r);
                    row.Write($"row {r} text");
                }
                await sheet.EndAsync(ct);
            }

            ms.Position = 0;
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            int rowIndex = 0;
            while (e.MoveNext())
            {
                Row row = e.Current;
                Assert.Equal(rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture), row[0].GetString());
                rowIndex++;
            }
            Assert.Equal(Rows, rowIndex);
        }

        // XLSB's offload path only ever engages once the record buffer actually spills past
        // SpillThreshold (see XlsbSheetWriter.EnsureStream) — a small sheet must still round-trip
        // correctly with prefetchWrite: true even though WriteOffloadStream is never constructed.
        [Fact]
        public async Task XlsbPrefetchWriteRoundTripsASmallSheetThatNeverSpills()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using MemoryStream ms = new();
            await using (XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, prefetchWrite: true, ct: ct))
            {
                await wb.StartAsync(ct);
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(ct);
                for (int r = 0; r < 5; r++)
                {
                    await using XlsbRowWriter row = await sheet.StartRowAsync(ct);
                    row.Write(r);
                    row.Write($"row {r}");
                }
                await sheet.EndAsync(ct);
            }

            ms.Position = 0;
            using XlsbReader reader = Excel.FromXlsb(ms);
            using XlsbReader.Enumerator e = reader.GetEnumerator();
            int rowIndex = 0;
            while (e.MoveNext())
            {
                Assert.Equal($"row {rowIndex}", e.Current[1].GetString());
                rowIndex++;
            }
            Assert.Equal(5, rowIndex);
        }

        private static async Task<byte[]> BuildXlsxAsync(bool prefetchWrite, CancellationToken ct)
        {
            MemoryStream ms = new();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true, prefetchWrite: prefetchWrite, ct: ct))
            {
                await wb.StartAsync(ct);
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(ct);
                for (int r = 0; r < Rows; r++)
                {
                    await using XlsxRowWriter row = sheet.StartRow(ct);
                    row.Write(r);
                    row.Write($"row {r} text");
                }
                await sheet.EndAsync(ct);
            }
            return ms.ToArray();
        }

        private static async Task<byte[]> BuildXlsbAsync(bool prefetchWrite, CancellationToken ct)
        {
            MemoryStream ms = new();
            await using (XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, prefetchWrite: prefetchWrite, ct: ct))
            {
                await wb.StartAsync(ct);
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(ct);
                for (int r = 0; r < Rows; r++)
                {
                    await using XlsbRowWriter row = await sheet.StartRowAsync(ct);
                    row.Write(r);
                    row.Write($"row {r} text");
                }
                await sheet.EndAsync(ct);
            }
            return ms.ToArray();
        }
    }
}
