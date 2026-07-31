using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    // CsvWorkbookWriter previously had no WriterState tracking at all — AddSheet worked
    // identically before StartAsync, after EndAsync, or called a second time, and had no null check.
    // These pin down the state machine now shared with XlsxWorkbookWriter/XlsbWorkbookWriter/XlsWorkbookWriter.
    public class CsvWorkbookWriterTests
    {
        [Fact]
        public async Task AddSheetBeforeStartThrows()
        {
            using var ms = new MemoryStream();
            await using var wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);

            Assert.Throws<InvalidOperationException>(() => wb.AddSheet("Sheet1"));
        }

        [Fact]
        public async Task AddSheetWithNullNameThrows()
        {
            using var ms = new MemoryStream();
            await using var wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);
            await wb.StartAsync(TestContext.Current.CancellationToken);

            Assert.Throws<ArgumentNullException>(() => wb.AddSheet(null!));
        }

        [Fact]
        public async Task SecondAddSheetThrows()
        {
            using var ms = new MemoryStream();
            await using var wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            wb.AddSheet("Sheet1");

            Assert.Throws<InvalidOperationException>(() => wb.AddSheet("Sheet2"));
        }

        [Fact]
        public async Task AddSheetAfterEndThrows()
        {
            using var ms = new MemoryStream();
            await using var wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            CsvSheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.EndAsync(TestContext.Current.CancellationToken);
            await wb.EndAsync(TestContext.Current.CancellationToken);

            Assert.Throws<ObjectDisposedException>(() => wb.AddSheet("Sheet2"));
        }

        [Fact]
        public async Task EndBeforeStartThrows()
        {
            using var ms = new MemoryStream();
            await using var wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);

            await Assert.ThrowsAsync<InvalidOperationException>(
                () => wb.EndAsync(TestContext.Current.CancellationToken).AsTask());
        }

        [Fact]
        public async Task StartTwiceThrows()
        {
            using var ms = new MemoryStream();
            await using var wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);
            await wb.StartAsync(TestContext.Current.CancellationToken);

            await Assert.ThrowsAsync<InvalidOperationException>(
                () => wb.StartAsync(TestContext.Current.CancellationToken).AsTask());
        }

        [Fact]
        public async Task NormalWriteRoundTripsThroughWorkbookRecordWriter()
        {
            using var ms = new MemoryStream();
            await using (var wb = CsvWorkbookWriter.Create(ms, leaveOpen: true))
            {
                await wb.StartAsync(TestContext.Current.CancellationToken);
                CsvSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (CsvRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write("hello");
                    row.Write(42);
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }

            ms.Position = 0;
            string text = new StreamReader(ms).ReadToEnd();
            Assert.Contains("hello", text, StringComparison.Ordinal);
            Assert.Contains("42", text, StringComparison.Ordinal);
        }
    }
}
