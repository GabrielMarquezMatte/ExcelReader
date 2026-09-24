using System.Globalization;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Tests.Reader
{
    public class FastDateTests
    {
        [Theory]
        [InlineData("2024-03-15")]
        [InlineData("2024-03-15T10:20:30")]
        [InlineData("2024-03-15 10:20:30")]
        [InlineData("2024-03-15T10:20:30.1")]
        [InlineData("2024-03-15T10:20:30.1234567")]
        [InlineData("2024-02-29")]
        [InlineData("0001-01-01")]
        [InlineData("9999-12-31T23:59:59.9999999")]
        public void AcceptsIsoAndMatchesTheGeneralParser(string text)
        {
            bool ok = FastDate.TryParse(Encoding.ASCII.GetBytes(text), out DateTime actual);

            Assert.True(ok, $"FastDate rejected \"{text}\".");
            Assert.True(DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime expected));
            Assert.Equal(expected, actual);
        }

        [Theory]
        [InlineData("")]
        [InlineData("not-a-date")]
        [InlineData("2024-03")]
        [InlineData("2024-3-15")]
        [InlineData("2024/03/15")]
        [InlineData("2023-02-29")]
        [InlineData("2024-13-01")]
        [InlineData("2024-00-10")]
        [InlineData("2024-03-00")]
        [InlineData("2024-03-32")]
        [InlineData("2024-03-15T24:00:00")]
        [InlineData("2024-03-15T10:60:00")]
        [InlineData("2024-03-15T10:20:60")]
        [InlineData("2024-03-15X10:20:30")]
        [InlineData("2024-03-15T10:20:30.12345678")]
        [InlineData("2024-03-15T10:20:30.")]
        [InlineData("2024-03-15T10:20")]
        public void Rejects(string text)
        {
            Assert.False(FastDate.TryParse(Encoding.ASCII.GetBytes(text), out _));
        }

        [Theory]
        [InlineData("2024503-15T10:20:30.1234567")]
        [InlineData("2024-03515T10:20:30.1234567")]
        [InlineData("2024-03-15T10520:30.1234567")]
        [InlineData("2024-03-15T10:20530.1234567")]
        [InlineData("2024-03-15T10:20:3051234567")]
        [InlineData("202X-03-15T10:20:30.1234567")]
        [InlineData("2024-03-15T10:20:30.123456X")]
        [InlineData("2024-02-30T10:20:30.1234567")]
        [InlineData("2024-03-15T24:20:30.1234567")]
        [InlineData("2024-03-15T10:60:30.1234567")]
        [InlineData("2024-03-15T10:20:60.1234567")]
        [InlineData("0000-03-15T10:20:30.1234567")]
        public void RejectsMalformedRoundTripForm(string text)
        {
            Assert.Equal(27, text.Length);
            Assert.False(FastDate.TryParse(Encoding.ASCII.GetBytes(text), out _));
        }

        [Fact]
        public void RoundTripFormAgreesWithTheGeneralParserAcrossADecade()
        {
            var cursor = new DateTime(2015, 1, 1, 0, 0, 0, DateTimeKind.Unspecified).AddTicks(1234567);
            for (int i = 0; i < 4000; i++)
            {
                string text = cursor.ToString("O", CultureInfo.InvariantCulture);
                Assert.Equal(27, text.Length);

                Assert.True(FastDate.TryParse(Encoding.ASCII.GetBytes(text), out DateTime actual), text);
                Assert.True(DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime expected));
                Assert.Equal(expected, actual);

                cursor = cursor.AddDays(1).AddHours(7).AddMinutes(13).AddSeconds(29).AddTicks(1111);
            }
        }

        [Fact]
        public void RoundTripFormAcceptsASpaceSeparator()
        {
            Assert.True(FastDate.TryParse("2024-03-15 10:20:30.1234567"u8, out DateTime actual));
            Assert.Equal(new DateTime(2024, 3, 15, 10, 20, 30).AddTicks(1234567), actual);
        }

        [Theory]
        [InlineData("2024-01-3/T00:00:00.0000000")]
        [InlineData("2024-01-3 T00:00:00.0000000")]
        public void RejectsNonDigitInDaySecondDigit(string text)
        {
            Assert.Equal(27, text.Length);
            Assert.False(FastDate.TryParse(Encoding.ASCII.GetBytes(text), out _));
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(2)]
        [InlineData(3)]
        [InlineData(5)]
        [InlineData(6)]
        [InlineData(8)]
        [InlineData(9)]
        [InlineData(11)]
        [InlineData(12)]
        [InlineData(14)]
        [InlineData(15)]
        public void RejectsNonDigitAtEveryDigitPosition(int position)
        {
            byte[] bytes = Encoding.ASCII.GetBytes("2024-03-15T10:20:30.1234567");
            bytes[position] = (byte)'/';

            Assert.False(FastDate.TryParse(bytes, out _));
        }

        [Theory]
        [InlineData("2024-03-15T10:20:30Z")]
        [InlineData("2024-03-15T10:20:30+01:00")]
        [InlineData("2024-03-15T10:20:30.500Z")]
        public void RejectsZoneDesignators(string text)
        {
            Assert.False(FastDate.TryParse(Encoding.ASCII.GetBytes(text), out _));
        }

        [Fact]
        public void CsvTextDatesAreReadableThroughTryGetDateTime()
        {
            byte[] csv = Encoding.UTF8.GetBytes("a,2024-03-15\nb,2024-03-16T07:08:09\n");
            using var ms = new MemoryStream(csv, writable: false);
            using CsvReader reader = Excel.FromCsv(ms);

            var dates = new List<DateTime>();
            foreach (Row row in reader)
            {
                int col = 0;
                foreach (RowCell rowCell in row.Cells)
                {
                    if (col == 1)
                    {
                        Assert.True(rowCell.Value.TryGetDateTime(out DateTime d));
                        dates.Add(d);
                    }
                    col++;
                }
            }

            Assert.Equal([new DateTime(2024, 3, 15), new DateTime(2024, 3, 16, 7, 8, 9)], dates);
        }

        [Fact]
        public void NumericCellsStillReadAsExcelSerials()
        {
            byte[] csv = Encoding.UTF8.GetBytes("a,45366\n");
            using var ms = new MemoryStream(csv, writable: false);
            using CsvReader reader = Excel.FromCsv(ms);

            foreach (Row row in reader)
            {
                Assert.True(row[1].TryGetDateTime(out DateTime d));
                Assert.Equal(new DateTime(2024, 3, 15), d);
            }
        }
    }
}
