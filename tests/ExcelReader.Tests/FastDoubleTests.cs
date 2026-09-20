using System.Globalization;
using System.Text;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    public class FastDoubleTests
    {
        private static void AssertMatchesDoubleTryParse(string text)
        {
            byte[] utf8 = Encoding.ASCII.GetBytes(text);
            bool expectedOk = double.TryParse(text, CultureInfo.InvariantCulture, out double expected);
            bool actualOk = FastDouble.TryParse(utf8, out double actual);

            if (!expectedOk)
            {
                Assert.False(actualOk, $"FastDouble accepted \"{text}\" but double.TryParse rejected it.");
                return;
            }
            if (actualOk)
            {
                Assert.Equal(expected, actual);
            }
        }

        [Theory]
        [InlineData("0")]
        [InlineData("1")]
        [InlineData("42")]
        [InlineData("-42")]
        [InlineData("+42")]
        [InlineData("3.14")]
        [InlineData("-3.14")]
        [InlineData("0.1")]
        [InlineData("100000")]
        [InlineData("999999999999999")]
        [InlineData("1.")]
        [InlineData(".5")]
        [InlineData("-.5")]
        [InlineData("0.0")]
        [InlineData("-0")]
        [InlineData("00042")]
        [InlineData("123456789.123456")]
        [InlineData("000000000000000123")]
        [InlineData("0.00000000000000123")]
        public void AcceptsAndMatchesPlainDecimals(string text)
        {
            AssertMatchesDoubleTryParse(text);
            byte[] utf8 = Encoding.ASCII.GetBytes(text);
            Assert.True(FastDouble.TryParse(utf8, out _), $"Expected FastDouble to accept \"{text}\".");
        }

        [Theory]
        [InlineData("1e5")]
        [InlineData("1E5")]
        [InlineData("1e-5")]
        [InlineData("1.5e10")]
        [InlineData("12345678901234567")]
        [InlineData("9999999999999999999999999")]
        [InlineData("")]
        [InlineData("-")]
        [InlineData("+")]
        [InlineData(".")]
        [InlineData("-.")]
        [InlineData("1.2.3")]
        [InlineData("1,000")]
        [InlineData("abc")]
        [InlineData("1x")]
        [InlineData("--5")]
        [InlineData("NaN")]
        [InlineData("Infinity")]
        public void RejectsUnsupportedForms(string text)
        {
            byte[] utf8 = Encoding.ASCII.GetBytes(text);
            Assert.False(FastDouble.TryParse(utf8, out _), $"Expected FastDouble to reject \"{text}\".");
            AssertMatchesDoubleTryParse(text);
        }

        [Fact]
        public void NegativeZeroMatchesSignBit()
        {
            Assert.True(FastDouble.TryParse("-0"u8, out double d));
            Assert.True(double.IsNegative(d));
        }

        [Theory]
        [InlineData(0.0)]
        [InlineData(1.0)]
        [InlineData(-1.0)]
        [InlineData(3.14)]
        [InlineData(0.5)]
        [InlineData(123456789.0)]
        [InlineData(-987654321.125)]
        public void RoundTripsFormattedDoubles(double value)
        {
            string text = value.ToString("G17", CultureInfo.InvariantCulture);
            AssertMatchesDoubleTryParse(text);
        }
    }
}
