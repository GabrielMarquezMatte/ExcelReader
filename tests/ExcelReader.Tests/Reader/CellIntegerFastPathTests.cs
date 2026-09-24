using System.Globalization;
using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests.Reader
{
    public class CellIntegerFastPathTests
    {
        private static Cell Text(string value)
        {
            return new Cell(CellType.ExcelString, Encoding.UTF8.GetBytes(value));
        }

        private static readonly CultureInfo Swedish = CultureInfo.GetCultureInfo("sv-SE");

        [Theory]
        [InlineData("0")]
        [InlineData("7")]
        [InlineData("123")]
        [InlineData("999999999")]
        [InlineData("1000000000")]
        [InlineData("2147483647")]
        [InlineData("2147483648")]
        [InlineData("007")]
        [InlineData("-1")]
        [InlineData("-2147483648")]
        [InlineData("-2147483649")]
        [InlineData("+5")]
        [InlineData(" 42")]
        [InlineData("42 ")]
        [InlineData("1,234")]
        [InlineData("1.5")]
        [InlineData("")]
        [InlineData("-")]
        [InlineData("12a")]
        [InlineData("a12")]
        [InlineData("１２３")]
        public void IntAgreesWithTheFrameworkParser(string raw)
        {
            byte[] utf8 = Encoding.UTF8.GetBytes(raw);
            bool expectedOk = int.TryParse(utf8, CultureInfo.InvariantCulture, out int expected);

            bool ok = Text(raw).TryParse(CultureInfo.InvariantCulture, out int actual);

            Assert.Equal(expectedOk, ok);
            Assert.Equal(expected, actual);
        }

        [Theory]
        [InlineData("0")]
        [InlineData("123")]
        [InlineData("999999999999999999")]
        [InlineData("1000000000000000000")]
        [InlineData("9223372036854775807")]
        [InlineData("9223372036854775808")]
        [InlineData("-9223372036854775808")]
        [InlineData("-1")]
        [InlineData("1.5")]
        [InlineData("")]
        public void LongAgreesWithTheFrameworkParser(string raw)
        {
            byte[] utf8 = Encoding.UTF8.GetBytes(raw);
            bool expectedOk = long.TryParse(utf8, CultureInfo.InvariantCulture, out long expected);

            bool ok = Text(raw).TryParse(CultureInfo.InvariantCulture, out long actual);

            Assert.Equal(expectedOk, ok);
            Assert.Equal(expected, actual);
        }

        [Theory]
        [InlineData("123")]
        [InlineData("-123")]
        [InlineData("−123")]
        public void ForeignCultureAgreesWithTheFrameworkParser(string raw)
        {
            byte[] utf8 = Encoding.UTF8.GetBytes(raw);
            bool expectedOk = int.TryParse(utf8, Swedish, out int expected);

            bool ok = Text(raw).TryParse(Swedish, out int actual);

            Assert.Equal(expectedOk, ok);
            Assert.Equal(expected, actual);
        }

        [Fact]
        public void NumberFormatInfoProviderStillParsesSignedValues()
        {
            bool ok = Text("-42").TryParse(NumberFormatInfo.InvariantInfo, out int value);

            Assert.True(ok);
            Assert.Equal(-42, value);
        }

        [Fact]
        public void AgreesWithTheFrameworkParserAcrossEveryDigitLength()
        {
            for (int digits = 1; digits <= 12; digits++)
            {
                foreach (string raw in new[] { new string('9', digits), "1" + new string('0', digits - 1) })
                {
                    byte[] utf8 = Encoding.UTF8.GetBytes(raw);
                    bool expectedOk = int.TryParse(utf8, CultureInfo.InvariantCulture, out int expected);

                    bool ok = Text(raw).TryParse(CultureInfo.InvariantCulture, out int actual);

                    Assert.Equal(expectedOk, ok);
                    Assert.Equal(expected, actual);
                }
            }
        }
    }
}
