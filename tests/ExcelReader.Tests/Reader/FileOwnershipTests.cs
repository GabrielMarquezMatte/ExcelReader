using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Tests.Reader.Xls;

namespace ExcelReader.Tests.Reader
{
    public sealed class FileOwnershipTests : IDisposable
    {
        private readonly string _path = Path.GetTempFileName();

        public void Dispose()
        {
            File.Delete(_path);
        }

        private static readonly CsvReaderOptions DelimiterEqualsQuote = new() { Delimiter = (byte)'"' };

        public static TheoryData<string> FailedOpens => ["FromXlsxFile", "FromXlsbFile", "FromCsvFile", "FromCsvFileAsync", "FromCsvFileAsync(canceled)"];

        [Theory]
        [MemberData(nameof(FailedOpens))]
        public async Task Should_ReleaseTheFile_When_OpeningItFails(string entryPoint)
        {
            File.WriteAllBytes(_path, entryPoint.StartsWith("FromXls", StringComparison.Ordinal)
                ? "PK\x03\x04 not a zip archive, just its signature"u8.ToArray()
                : "a,b\n1,2\n"u8.ToArray());

            Exception? thrown = await Record.ExceptionAsync(() => Open(entryPoint));

            Assert.NotNull(thrown);
            using var exclusive = new FileStream(_path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }

        [Fact]
        public async Task Should_HonorACanceledToken_When_OpeningAValidXlsFileAsync()
        {
            File.WriteAllBytes(_path, ValidXls());

            await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => Excel.FromXlsFileAsync(_path, ct: new CancellationToken(canceled: true)).AsTask());

            using var exclusive = new FileStream(_path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task Should_ReleaseAnOwnedNonSeekableXlsStream_When_BufferingItFails(bool async)
        {
            var stream = new NonSeekableStream(ValidXls()) { OnRead = static () => throw new IOException("read failed") };

            if (async)
            {
                await Assert.ThrowsAsync<IOException>(
                    () => Excel.FromXlsAsync(stream, leaveOpen: false, ct: TestContext.Current.CancellationToken).AsTask());
            }
            else
            {
                Assert.Throws<IOException>(() => Excel.FromXls(stream, leaveOpen: false));
            }

            Assert.True(stream.Disposed);
        }

        private static byte[] ValidXls()
        {
            using MemoryStream built = XlsWorkbookBuilder.Build(sheets: [("S1", [["Name", 1, true]])]);
            return built.ToArray();
        }

        private async Task Open(string entryPoint)
        {
            switch (entryPoint)
            {
                case "FromXlsxFile":
                    Excel.FromXlsxFile(_path).Dispose();
                    break;
                case "FromXlsbFile":
                    Excel.FromXlsbFile(_path).Dispose();
                    break;
                case "FromCsvFile":
                    Excel.FromCsvFile(_path, DelimiterEqualsQuote).Dispose();
                    break;
                case "FromCsvFileAsync":
                    (await Excel.FromCsvFileAsync(_path, DelimiterEqualsQuote)).Dispose();
                    break;
                default:
                    (await Excel.FromCsvFileAsync(_path, ct: new CancellationToken(canceled: true))).Dispose();
                    break;
            }
        }
    }
}
