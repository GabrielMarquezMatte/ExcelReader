using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Writer.Csv;

namespace ExcelReader.Tests.Writer.Csv
{
    public class CsvWriterOptionsTests
    {
        private static byte[] Write(CsvWriterOptions options)
        {
            var stream = new MemoryStream();
            using (CsvWriter writer = CsvWriter.Create(stream, leaveOpen: true, options))
            {
                using (CsvRowWriter header = writer.StartRow())
                {
                    header.Write("Name");
                    header.Write("City");
                }
                using (CsvRowWriter row = writer.StartRow())
                {
                    row.Write("José");
                    row.Write("São Paulo");
                }
            }
            return stream.ToArray();
        }

        [Fact]
        public void Should_TerminateRecordsWithLineFeed_When_NewLineIsLineFeed()
        {
            byte[] written = Write(CsvWriterOptions.Default with { NewLine = CsvNewLine.LineFeed });

            Assert.DoesNotContain((byte)'\r', written);
            Assert.Equal("Name,City\nJosé,São Paulo\n", Encoding.UTF8.GetString(written));
        }

        [Fact]
        public void Should_TerminateRecordsWithCarriageReturnLineFeed_When_NewLineIsDefaulted()
        {
            byte[] written = Write(CsvWriterOptions.Default);

            Assert.Equal("Name,City\r\nJosé,São Paulo\r\n", Encoding.UTF8.GetString(written));
        }

        [Fact]
        public void Should_WriteAByteOrderMark_When_Requested()
        {
            byte[] written = Write(CsvWriterOptions.Default with { WriteByteOrderMark = true });

            Assert.Equal(Encoding.UTF8.GetPreamble(), written[..3]);
            Assert.StartsWith("Name,City", Encoding.UTF8.GetString(written.AsSpan(3)), StringComparison.Ordinal);
        }

        [Fact]
        public void Should_TranscodeOutput_When_EncodingIsSet()
        {
            byte[] written = Write(CsvWriterOptions.Default with { Encoding = Encoding.Latin1 });

            Assert.Equal("Name,City\r\nJosé,São Paulo\r\n", Encoding.Latin1.GetString(written));
            Assert.DoesNotContain((byte)0xC3, written);
        }

        [Fact]
        public void Should_WriteTheEncodingsPreamble_When_TranscodingWithAByteOrderMark()
        {
            byte[] written = Write(CsvWriterOptions.Default with { Encoding = Encoding.Unicode, WriteByteOrderMark = true });

            Assert.Equal(Encoding.Unicode.GetPreamble(), written[..2]);
        }

        [Fact]
        public void Should_RoundTripThroughTheReader_When_BothAreGivenTheSameDialect()
        {
            byte[] written = Write(CsvWriterOptions.Default with
            {
                Delimiter = (byte)';',
                Encoding = Encoding.Latin1,
                NewLine = CsvNewLine.LineFeed,
            });

            var options = CsvReaderOptions.Default with { Delimiter = (byte)';', Encoding = Encoding.Latin1 };
            using CsvReader reader = Excel.FromCsv(written.AsMemory(), options);
            using CsvReader.Enumerator rows = reader.GetEnumerator();

            Assert.True(rows.MoveNext());
            Assert.True(rows.MoveNext());
            var cells = new List<string>();
            foreach (RowCell cell in rows.Current.Cells)
            {
                cells.Add(cell.Value.GetString());
            }
            Assert.Equal(["José", "São Paulo"], cells);
        }

        [Fact]
        public void Should_Reject_When_NewLineIsNotADefinedValue()
        {
            var stream = new MemoryStream();

            Assert.Throws<ArgumentException>(
                () => CsvWriter.Create(stream, leaveOpen: true, CsvWriterOptions.Default with { NewLine = (CsvNewLine)7 }));
        }
    }
}
