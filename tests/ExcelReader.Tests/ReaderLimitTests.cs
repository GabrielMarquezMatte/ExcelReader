using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class ReaderLimitTests
    {
        [Fact]
        public void CompressedSheetTripsTotalDecompressedLimit()
        {
            string value = new('A', 256 * 1024);
            using MemoryStream ms = WorkbookBuilder.Build(
                $"""<row r="1"><c r="A1" t="inlineStr"><is><t>{value}</t></is></c></row>""");
            Assert.True(ms.Length < 10_000);

            var options = new ExcelReaderOptions
            {
                MaxTotalDecompressedBytes = 16 * 1024,
            };

            using XlsxReader reader = Excel.From(ms, options: options);
            ExcelLimitExceededException ex = Assert.Throws<ExcelLimitExceededException>(() =>
            {
                using XlsxReader.Enumerator e = reader.GetEnumerator();
                Assert.True(e.MoveNext());
            });
            Assert.Equal(nameof(ExcelReaderOptions.MaxTotalDecompressedBytes), ex.LimitName);
        }

        [Fact]
        public void SharedStringsTripSharedStringLimit()
        {
            string value = new('B', 128 * 1024);
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v>0</v></c></row>""",
                sharedStrings: $"<si><t>{value}</t></si>");
            Assert.True(ms.Length < 10_000);

            var options = new ExcelReaderOptions
            {
                MaxTotalDecompressedBytes = 0,
                MaxSharedStringBytes = 1024,
            };

            using XlsxReader reader = Excel.From(ms, options: options);
            ExcelLimitExceededException ex = Assert.Throws<ExcelLimitExceededException>(() =>
            {
                using XlsxReader.Enumerator e = reader.GetEnumerator();
            });
            Assert.Equal(nameof(ExcelReaderOptions.MaxSharedStringBytes), ex.LimitName);
        }

        [Fact]
        public void ZeroTotalDecompressedLimitRestoresUnlimitedTotalRead()
        {
            string value = new('C', 96 * 1024);
            using MemoryStream ms = WorkbookBuilder.Build(
                $"""<row r="1"><c r="A1" t="inlineStr"><is><t>{value}</t></is></c></row>""");

            var options = new ExcelReaderOptions
            {
                MaxTotalDecompressedBytes = 0,
            };

            using XlsxReader reader = Excel.From(ms, options: options);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(value.Length, e.Current[0].GetString().Length);
        }

        [Fact]
        public async Task LimitedReadStreamCoversUnsupportedOperationsAndByteArrayAsyncRead()
        {
            await using MemoryStream inner = new([1, 2, 3, 4, 5]);
            await using LimitedReadStream limited = new(inner, new DecompressedByteCounter(limit: 10));

            Assert.True(limited.CanRead);
            Assert.False(limited.CanSeek);
            Assert.False(limited.CanWrite);
            Assert.Throws<NotSupportedException>(() => limited.Length);
            Assert.Throws<NotSupportedException>(() => limited.Position);
            Assert.Throws<NotSupportedException>(() => limited.Position = 0);
            Assert.Throws<NotSupportedException>(() => limited.Seek(0, SeekOrigin.Begin));
            Assert.Throws<NotSupportedException>(() => limited.SetLength(0));
            Assert.Throws<NotSupportedException>(() => limited.Write([1], 0, 1));

            byte[] buffer = new byte[3];
            int read = await limited.ReadAsync(buffer, 0, buffer.Length, TestContext.Current.CancellationToken);
            Assert.Equal(3, read);
            Assert.Equal([1, 2, 3], buffer);

            limited.Flush();
        }

        [Fact]
        public void LimitedReadStreamCountsSpanReadsAndEntryLimit()
        {
            using MemoryStream inner = new([1, 2, 3, 4]);
            using LimitedReadStream limited = new(
                inner,
                totalCounter: null,
                entryLimitName: nameof(ExcelReaderOptions.MaxCellBytes),
                entryLimit: 2);

            Span<byte> buffer = stackalloc byte[2];
            Assert.Equal(2, limited.Read(buffer));

            ExcelLimitExceededException ex;
            try
            {
                limited.Read(buffer);
                throw new Xunit.Sdk.XunitException("Expected ExcelLimitExceededException.");
            }
            catch (ExcelLimitExceededException caught)
            {
                ex = caught;
            }

            Assert.Equal(nameof(ExcelReaderOptions.MaxCellBytes), ex.LimitName);
            Assert.Equal(2, ex.Limit);
            Assert.Equal(4, ex.Actual);
        }

        [Fact]
        public void ExcelLimitExceededExceptionConstructorsAreCovered()
        {
            var empty = new ExcelLimitExceededException();
            Assert.Equal(string.Empty, empty.LimitName);

            var withMessage = new ExcelLimitExceededException("message");
            Assert.Equal("message", withMessage.Message);

            var inner = new InvalidOperationException("inner");
            var withInner = new ExcelLimitExceededException("outer", inner);
            Assert.Equal("outer", withInner.Message);
            Assert.Same(inner, withInner.InnerException);
        }
    }
}
