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
    }
}
