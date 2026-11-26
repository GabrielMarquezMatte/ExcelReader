using ExcelReader.Core.Parser;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer
{
    public class WriteRecordsFailureTests
    {
        private sealed class Row
        {
            [ExcelColumn("Data")]
            public DateOnly Date { get; set; }
        }

        private static CancellationToken Ct => TestContext.Current.CancellationToken;

        private static async IAsyncEnumerable<int> OneAsync()
        {
            await Task.Yield();
            yield return 1;
        }

        [Fact]
        public async Task LayoutWriteFailureReachesTheCallerThroughSheetDispose()
        {
            await Assert.ThrowsAsync<ArgumentOutOfRangeException>(async () =>
            {
                using var stream = new MemoryStream();
                await using var writer = XlsxWorkbookWriter.Create(stream, leaveOpen: true);
                await using var sheet = writer.AddSheet("S");
                await sheet.WriteRecordsAsync([new Row()], ExcelRecordLayout.FromAttributes<Row>(), Ct);
            });
        }

        [Fact]
        public async Task XlsxFastPathKeepsTheOriginalFailureEvenWithAPendingStyle()
        {
            await Assert.ThrowsAsync<FormatException>(async () =>
            {
                using var stream = new MemoryStream();
                await using var writer = XlsxWorkbookWriter.Create(stream, leaveOpen: true);
                await using var sheet = writer.AddSheet("S");
                await sheet.WriteRecordsAsync([1], static (XlsxRowWriter row, int _) =>
                {
                    row.Style(0);
                    throw new FormatException();
                }, Ct);
            });
        }

        [Fact]
        public async Task XlsbWriteFailureReachesTheCaller()
        {
            await Assert.ThrowsAsync<FormatException>(async () =>
            {
                using var stream = new MemoryStream();
                await using var writer = XlsbWorkbookWriter.Create(stream, leaveOpen: true);
                await using var sheet = writer.AddSheet("S");
                await sheet.WriteRecordsAsync([1], static (XlsbRowWriter _, int _) => throw new FormatException(), Ct);
            });
        }

        [Fact]
        public async Task XlsAsyncSourceWriteFailureReachesTheCaller()
        {
            await Assert.ThrowsAsync<FormatException>(async () =>
            {
                using var stream = new MemoryStream();
                await using var writer = XlsWorkbookWriter.Create(stream, leaveOpen: true);
                await using var sheet = writer.AddSheet("S");
                await sheet.WriteRecordsAsync(OneAsync(), static (XlsRowWriter _, int _) => throw new FormatException(), Ct);
            });
        }
    }
}
