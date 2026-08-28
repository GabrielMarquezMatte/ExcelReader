using System.Globalization;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    // Cell.TryParse has an ASCII-digit shortcut for int/long that bypasses the framework number
    // parser. Its whole justification is that it cannot disagree with the parser it shortcuts, so
    // every case here asserts exactly that: same bool, same value, against the general parser it is
    // standing in for — including the shapes it must decline (signs under a foreign culture,
    // whitespace, separators, overflow) rather than accept.
    public class CellIntegerFastPathTests
    {
        private static Cell Text(string value)
        {
            return new Cell(CellType.ExcelString, Encoding.UTF8.GetBytes(value));
        }

        private static readonly CultureInfo Swedish = CultureInfo.GetCultureInfo("sv-SE");

        [Theory]
        // Plain digits, the shape the shortcut exists for.
        [InlineData("0")]
        [InlineData("7")]
        [InlineData("123")]
        [InlineData("999999999")]
        // One digit past the shortcut's own int limit, so the fallback has to carry it.
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
        [InlineData("１２３")] // full-width digits: not ASCII, must fall through
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

        // A culture is free to spell its negative sign with something other than U+002D, so the
        // shortcut declines a leading '-' unless the provider is the invariant culture. Unsigned
        // digits stay on the fast path everywhere, because no culture can read them differently.
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

        // A NumberFormatInfo is not the invariant CultureInfo by reference, so a signed value must
        // still land on the fallback and still come out right.
        [Fact]
        public void NumberFormatInfoProviderStillParsesSignedValues()
        {
            bool ok = Text("-42").TryParse(NumberFormatInfo.InvariantInfo, out int value);

            Assert.True(ok);
            Assert.Equal(-42, value);
        }

        // Exhaustive sweep over the digit-count boundaries the shortcut branches on, rather than a
        // handful of spot values.
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
