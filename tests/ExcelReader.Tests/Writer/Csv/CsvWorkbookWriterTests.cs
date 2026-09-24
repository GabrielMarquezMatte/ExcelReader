using ExcelReader.Core.Writer.Csv;

namespace ExcelReader.Tests.Writer.Csv
{
    public class CsvWorkbookWriterTests
    {
        [Fact]
        public async Task AddSheetWithNullNameThrows()
        {
            using var ms = new MemoryStream();
            await using var wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);

            Assert.Throws<ArgumentNullException>(() => wb.AddSheet(null!));
        }

        [Fact]
        public async Task SecondAddSheetThrows()
        {
            using var ms = new MemoryStream();
            await using var wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);
            wb.AddSheet("Sheet1");

            Assert.Throws<InvalidOperationException>(() => wb.AddSheet("Sheet2"));
        }

        [Fact]
        public async Task AddSheetAfterEndThrows()
        {
            using var ms = new MemoryStream();
            await using var wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);
            CsvSheetWriter sheet = wb.AddSheet("Sheet1");
            await sheet.EndAsync(TestContext.Current.CancellationToken);
            await wb.EndAsync(TestContext.Current.CancellationToken);

            Assert.Throws<ObjectDisposedException>(() => wb.AddSheet("Sheet2"));
        }

        [Fact]
        public async Task SheetRejectsRowsAndEndAfterEnd()
        {
            using var ms = new MemoryStream();
            await using var wb = CsvWorkbookWriter.Create(ms, leaveOpen: true);
            CsvSheetWriter sheet = wb.AddSheet("Sheet1");
            sheet.End();

            Assert.Throws<ObjectDisposedException>(() => sheet.StartRow());
            await Assert.ThrowsAsync<ObjectDisposedException>(async () => await sheet.StartRowAsync(TestContext.Current.CancellationToken));
            Assert.Throws<ObjectDisposedException>(sheet.End);
        }

        [Fact]
        public async Task NormalWriteRoundTripsThroughWorkbookRecordWriter()
        {
            using var ms = new MemoryStream();
            await using (var wb = CsvWorkbookWriter.Create(ms, leaveOpen: true))
            {
                CsvSheetWriter sheet = wb.AddSheet("Sheet1");
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
