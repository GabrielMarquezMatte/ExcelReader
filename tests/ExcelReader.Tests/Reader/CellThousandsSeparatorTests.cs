using System.Globalization;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests.Reader
{
    public class CellThousandsSeparatorTests
    {
        private static readonly CultureInfo Brazilian = CultureInfo.GetCultureInfo("pt-BR");

        [Theory]
        [InlineData("1,5")]
        [InlineData("1,000.5")]
        [InlineData("1,234")]
        public void InvariantTextWithAThousandsSeparatorIsNotANumber(string raw)
        {
            var cell = new Cell(CellType.ExcelString, System.Text.Encoding.ASCII.GetBytes(raw));

            Assert.False(cell.TryParse(CultureInfo.InvariantCulture, out double _));
            Assert.False(cell.TryParse(CultureInfo.InvariantCulture, out float _));
            Assert.False(cell.TryParse(CultureInfo.InvariantCulture, out decimal _));
            Assert.False(cell.TryParse(CultureInfo.InvariantCulture, out Half _));
            Assert.False(cell.TryGetDouble(out _));
        }

        [Fact]
        public void DecimalCommaStillParsesUnderACommaCulture()
        {
            var cell = new Cell(CellType.ExcelString, "1,5"u8);

            Assert.True(cell.TryParse(Brazilian, out double d));
            Assert.Equal(1.5, d);
            Assert.True(cell.TryParse(Brazilian, out decimal m));
            Assert.Equal(1.5m, m);
        }

        [Fact]
        public void ExponentsStillParse()
        {
            var cell = new Cell(CellType.ExcelString, "1.5E3"u8);

            Assert.True(cell.TryParse(CultureInfo.InvariantCulture, out double d));
            Assert.Equal(1500.0, d);
            Assert.True(cell.TryParse(CultureInfo.InvariantCulture, out decimal m));
            Assert.Equal(1500m, m);
        }
    }
}
