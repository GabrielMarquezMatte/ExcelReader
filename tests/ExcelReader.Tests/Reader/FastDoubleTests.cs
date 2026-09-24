using System.Globalization;
using System.Text;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Tests.Reader
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
        [InlineData("1e+5")]
        [InlineData("1e-5")]
        [InlineData("1e0")]
        [InlineData("1.5e10")]
        [InlineData("0.5e1")]
        [InlineData("9.99E2")]
        [InlineData("-2.25E-7")]
        [InlineData("6.022e23")]
        [InlineData("-1.6e-19")]
        [InlineData("1e22")]
        [InlineData("1e-22")]
        [InlineData("-0e5")]
        public void AcceptsAndMatchesScientificNotation(string text)
        {
            AssertMatchesDoubleTryParse(text);
            byte[] utf8 = Encoding.ASCII.GetBytes(text);
            Assert.True(FastDouble.TryParse(utf8, out _), $"Expected FastDouble to accept \"{text}\".");
        }

        [Theory]
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
        [InlineData("1e")]
        [InlineData("1e+")]
        [InlineData("1e-")]
        [InlineData("e5")]
        [InlineData("1ee5")]
        [InlineData("1e+-5")]
        [InlineData("1e5.5")]
        [InlineData("1e1000")]
        [InlineData("1e23")]
        [InlineData("1e309")]
        [InlineData("1e-400")]
        [InlineData("5E-324")]
        [InlineData("1.7976931348623157E+308")]
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
