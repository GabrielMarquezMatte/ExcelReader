using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    // Cancelling partway through a large sheet, not just at open time (PrefetchDecompressionTests
    // already covers the already-cancelled/prefetch cases). XLS and XLSB check the token on every
    // MoveNextAsync call, so cancellation is observed on the very next call regardless of buffering.
    // XLSX and CSV only observe it when the pooled buffer actually needs a refill (via the
    // underlying Stream.ReadAsync honoring the token), so those two need a workbook big enough to
    // force a real refill after cancellation to prove the token is wired through at all.
    public class MidStreamCancellationTests
    {
        [Fact]
        public async Task XlsxCancellationMidLargeSheetThrows()
        {
            CancellationToken outer = TestContext.Current.CancellationToken;
            byte[] bytes = BuildLargeXlsx();
            using CancellationTokenSource cts = new();

            await using MemoryStream ms = new(bytes, writable: false);
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: outer);
            await using XlsxReader.Enumerator e = await reader.GetAsyncEnumeratorAsync(cts.Token);

            for (int i = 0; i < 5; i++)
            {
                Assert.True(await e.MoveNextAsync());
            }
            cts.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            {
                while (await e.MoveNextAsync())
                {
                    Assert.True(e.Current.ColumnCount >= 0);
                }
            });
        }

        [Fact]
        public async Task CsvCancellationMidLargeSheetThrows()
        {
            CancellationToken outer = TestContext.Current.CancellationToken;
            byte[] bytes = BuildLargeCsv();
            using CancellationTokenSource cts = new();

            await using MemoryStream ms = new(bytes, writable: false);
            await using CsvReader reader = await Excel.FromCsvAsync(ms, ct: outer);
            await using CsvReader.Enumerator e = await reader.GetAsyncEnumeratorAsync(cts.Token);

            for (int i = 0; i < 5; i++)
            {
                Assert.True(await e.MoveNextAsync());
            }
            cts.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            {
                while (await e.MoveNextAsync())
                {
                    Assert.True(e.Current.ColumnCount >= 0);
                }
            });
        }

        [Fact]
        public async Task XlsCancellationAfterCancelThrowsOnNextMoveNext()
        {
            // XlsReader.Enumerator checks the token unconditionally on every call, so a small
            // workbook already proves the contract — no need to force a buffer refill.
            CancellationToken outer = TestContext.Current.CancellationToken;
            using MemoryStream ms = XlsWorkbookBuilder.Build(sheets: [("S1", [["Row1"], ["Row2"], ["Row3"]])]);
            using CancellationTokenSource cts = new();

            await using XlsReader reader = await Excel.FromXlsAsync(ms, ct: outer);
            await using XlsReader.Enumerator e = reader.GetAsyncEnumerator(cts.Token);

            Assert.True(await e.MoveNextAsync());
            cts.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await e.MoveNextAsync());
        }

        [Fact]
        public async Task XlsbCancellationAfterCancelThrowsOnNextMoveNext()
        {
            // XlsbReader.Enumerator checks the token unconditionally on every call too.
            CancellationToken outer = TestContext.Current.CancellationToken;
            byte[] bytes = await BuildSmallXlsbAsync(outer);
            using CancellationTokenSource cts = new();

            await using MemoryStream ms = new(bytes, writable: false);
            await using XlsbReader reader = await Excel.FromXlsbAsync(ms, ct: outer);
            await using XlsbReader.Enumerator e = await reader.GetAsyncEnumeratorAsync(cts.Token);

            Assert.True(await e.MoveNextAsync());
            cts.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await e.MoveNextAsync());
        }

        // Many mixed-type rows so the sheet spans several 64 KiB buffer refills, not just the
        // initial fill — the point where XLSX/CSV's implicit (stream-driven) cancellation fires.
        private static byte[] BuildLargeXlsx()
        {
            StringBuilder sb = new(512 * 1024);
            for (int r = 1; r <= 6000; r++)
            {
                sb.Append("<row r=\"").Append(r).Append("\">")
                  .Append("<c r=\"A").Append(r).Append("\"><v>").Append(r).Append("</v></c>")
                  .Append("<c r=\"B").Append(r).Append("\" t=\"inlineStr\"><is><t>row ").Append(r).Append(" text</t></is></c>")
                  .Append("</row>");
            }
            using MemoryStream ms = WorkbookBuilder.Build(sb.ToString());
            return ms.ToArray();
        }

        private static byte[] BuildLargeCsv()
        {
            StringBuilder sb = new(512 * 1024);
            for (int r = 1; r <= 20_000; r++)
            {
                sb.Append(r).Append(",row ").Append(r).Append(" text,").Append(r * 1.5).Append('\n');
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        private static async Task<byte[]> BuildSmallXlsbAsync(CancellationToken ct)
        {
            MemoryStream ms = new();
            await using (XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: ct))
            {
                await wb.StartAsync(ct);
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(ct);
                for (int r = 0; r < 3; r++)
                {
                    await using XlsbRowWriter row = await sheet.StartRowAsync(ct);
                    row.Write($"row {r}");
                }
                await sheet.EndAsync(ct);
                await wb.EndAsync(ct);
            }
            return ms.ToArray();
        }
    }
}
