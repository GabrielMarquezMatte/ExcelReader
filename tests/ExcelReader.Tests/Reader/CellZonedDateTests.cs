using ExcelReader.Core.Reader;

namespace ExcelReader.Tests.Reader
{
    public class CellZonedDateTests
    {
        [Theory]
        [InlineData("2024-01-01T10:00:00Z", 10)]
        [InlineData("2024-01-01T10:00:00+05:00", 5)]
        [InlineData("2024-01-01 10:00:00-03:00", 13)]
        [InlineData("2024-01-01T10:00:00.1234567Z", 10)]
        [InlineData("2024-01-01T10:00+05:00", 5)]
        public void ZonedIsoTextConvertsToUtc(string raw, int utcHour)
        {
            var cell = new Cell(CellType.ExcelString, System.Text.Encoding.ASCII.GetBytes(raw));

            Assert.True(cell.TryGetDateTime(isDate1904: false, out DateTime value));
            Assert.Equal(DateTimeKind.Utc, value.Kind);
            Assert.Equal(new DateTime(2024, 1, 1), value.Date);
            Assert.Equal(utcHour, value.Hour);
            Assert.Equal(0, value.Minute);
        }

        [Fact]
        public void UnzonedIsoTextKeepsItsClockTime()
        {
            var cell = new Cell(CellType.ExcelString, "2024-01-01T10:00:00"u8);

            Assert.True(cell.TryGetDateTime(isDate1904: false, out DateTime value));
            Assert.Equal(new DateTime(2024, 1, 1, 10, 0, 0), value);
            Assert.Equal(DateTimeKind.Unspecified, value.Kind);
        }

        [Theory]
        [InlineData("not a date")]
        [InlineData("01/02/2024 10:00Z")]
        public void NonIsoTextIsStillRejected(string raw)
        {
            var cell = new Cell(CellType.ExcelString, System.Text.Encoding.ASCII.GetBytes(raw));

            Assert.False(cell.TryGetDateTime(isDate1904: false, out _));
        }
    }
}
