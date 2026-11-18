using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Xls;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Tests.Reader.Xls;

namespace ExcelReader.Tests.Reader
{
    public class MidStreamCancellationTests
    {
        [Fact]
        public async Task XlsxCancellationMidLargeSheetThrows()
        {
            CancellationToken outer = TestContext.Current.CancellationToken;
            byte[] bytes = BuildLargeXlsx();
            using CancellationTokenSource cts = new();

            await using MemoryStream ms = new(bytes, writable: false);
            await using XlsxWorkbook reader = await Excel.FromXlsxAsync(ms, ct: outer);
            await using XlsxWorkbook.Enumerator e = reader.FirstSheet.GetAsyncEnumerator(cts.Token);

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
            await using CsvReader.Enumerator e = reader.FirstSheet.GetAsyncEnumerator(cts.Token);

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
        public async Task CsvCancellationAfterCancelThrowsOnNextMoveNext()
        {
            CancellationToken outer = TestContext.Current.CancellationToken;
            using CancellationTokenSource cts = new();

            await using MemoryStream ms = new("a,1\nb,2\nc,3\n"u8.ToArray(), writable: false);
            await using CsvReader reader = await Excel.FromCsvAsync(ms, ct: outer);
            await using CsvReader.Enumerator e = reader.FirstSheet.GetAsyncEnumerator(cts.Token);

            Assert.True(await e.MoveNextAsync());
            cts.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await e.MoveNextAsync());
        }

        [Fact]
        public async Task XlsCancellationAfterCancelThrowsOnNextMoveNext()
        {
            CancellationToken outer = TestContext.Current.CancellationToken;
            using MemoryStream ms = XlsWorkbookBuilder.Build(sheets: [("S1", [["Row1"], ["Row2"], ["Row3"]])]);
            using CancellationTokenSource cts = new();

            await using XlsWorkbook reader = await Excel.FromXlsAsync(ms, ct: outer);
            await using XlsWorkbook.Enumerator e = reader.FirstSheet.GetAsyncEnumerator(cts.Token);

            Assert.True(await e.MoveNextAsync());
            cts.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await e.MoveNextAsync());
        }

        [Fact]
        public async Task XlsxCancellationAfterCancelThrowsOnNextMoveNext()
        {
            CancellationToken outer = TestContext.Current.CancellationToken;
            using MemoryStream built = WorkbookBuilder.Build(
                "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"2\"><c r=\"A2\"><v>2</v></c></row><row r=\"3\"><c r=\"A3\"><v>3</v></c></row>");
            byte[] bytes = built.ToArray();
            using CancellationTokenSource cts = new();

            await using MemoryStream ms = new(bytes, writable: false);
            await using XlsxWorkbook reader = await Excel.FromXlsxAsync(ms, ct: outer);
            await using XlsxWorkbook.Enumerator e = reader.FirstSheet.GetAsyncEnumerator(cts.Token);

            Assert.True(await e.MoveNextAsync());
            cts.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await e.MoveNextAsync());
        }

        [Fact]
        public async Task XlsbCancellationAfterCancelThrowsOnNextMoveNext()
        {
            CancellationToken outer = TestContext.Current.CancellationToken;
            byte[] bytes = await BuildSmallXlsbAsync(outer);
            using CancellationTokenSource cts = new();

            await using MemoryStream ms = new(bytes, writable: false);
            await using XlsbWorkbook reader = await Excel.FromXlsbAsync(ms, ct: outer);
            await using XlsbWorkbook.Enumerator e = reader.FirstSheet.GetAsyncEnumerator(cts.Token);

            Assert.True(await e.MoveNextAsync());
            cts.Cancel();

            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await e.MoveNextAsync());
        }

        [Fact]
        public async Task XlsxMemoryBackedEnumeratorHonorsCancellation()
        {
            using MemoryStream built = WorkbookBuilder.Build("<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>");
            using CancellationTokenSource cts = new();
            cts.Cancel();

            await using XlsxWorkbook fromMemory = Excel.FromXlsx(built.ToArray());
            await using XlsxWorkbook fromWritableStream = await Excel.FromXlsxAsync(new MemoryStream(built.ToArray()), ct: TestContext.Current.CancellationToken);

            await AssertCancelsAsync(fromMemory.FirstSheet.GetAsyncEnumerator(cts.Token));
            await AssertCancelsAsync(fromWritableStream.FirstSheet.GetAsyncEnumerator(cts.Token));
        }

        [Fact]
        public async Task XlsbMemoryBackedEnumeratorHonorsCancellation()
        {
            byte[] bytes = await BuildSmallXlsbAsync(TestContext.Current.CancellationToken);
            using CancellationTokenSource cts = new();
            cts.Cancel();

            await using XlsbWorkbook fromMemory = Excel.FromXlsb(bytes);
            await using XlsbWorkbook fromWritableStream = await Excel.FromXlsbAsync(new MemoryStream(bytes), ct: TestContext.Current.CancellationToken);

            await AssertCancelsAsync(fromMemory.FirstSheet.GetAsyncEnumerator(cts.Token));
            await AssertCancelsAsync(fromWritableStream.FirstSheet.GetAsyncEnumerator(cts.Token));
        }

        private static async Task AssertCancelsAsync(IExcelRowEnumerator e)
        {
            await using (e)
            {
                await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await e.MoveNextAsync());
            }
        }

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
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
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
