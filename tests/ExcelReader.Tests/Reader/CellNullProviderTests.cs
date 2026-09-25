using System.Globalization;
using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests.Reader
{
    public class CellNullProviderTests
    {
        private static readonly CultureInfo Brazilian = CultureInfo.GetCultureInfo("pt-BR");

        private static T ParseUnderBrazilianCulture<T>(Cell cell) where T : struct, IUtf8SpanParsable<T>
        {
            CultureInfo previous = CultureInfo.CurrentCulture;
            CultureInfo.CurrentCulture = Brazilian;
            try
            {
                Assert.True(cell.TryParse(null, out T value));
                return value;
            }
            finally
            {
                CultureInfo.CurrentCulture = previous;
            }
        }

        [Theory]
        [InlineData("3.5")]
        [InlineData("35.840000000000003")]
        [InlineData("1.2345678901234567E-70")]
        [InlineData("12345678901234567890.5")]
        public void NullProviderParsesDoublesInvariantlyWhateverTheCurrentCulture(string raw)
        {
            var cell = new Cell(CellType.Number, Encoding.ASCII.GetBytes(raw));

            Assert.Equal(double.Parse(raw, CultureInfo.InvariantCulture), ParseUnderBrazilianCulture<double>(cell));
        }

        [Fact]
        public void NullProviderParsesDecimalsInvariantlyWhateverTheCurrentCulture()
        {
            var cell = new Cell(CellType.Number, "35.840000000000003"u8);

            Assert.Equal(35.840000000000003m, ParseUnderBrazilianCulture<decimal>(cell));
        }

        [Fact]
        public void StoredNumberFallbackParsesItsInvariantTextWhateverTheProvider()
        {
            var cell = new Cell(CellType.Number, "3.5"u8, 3.5, hasNumber: true, styleIndex: 0);

            Assert.True(cell.TryParse(Brazilian, out Half value));
            Assert.Equal((Half)3.5, value);
        }
    }
}
