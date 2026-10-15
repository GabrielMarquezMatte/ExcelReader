using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Tests.Reader
{
    public class FastDecimalTests
    {
        private static void AssertBitEqualToDecimalParse(string text)
        {
            byte[] utf8 = Encoding.ASCII.GetBytes(text);
            Assert.True(FastDecimal.TryParse(utf8, out decimal actual), $"Expected FastDecimal to accept \"{text}\".");
            decimal expected = decimal.Parse(utf8, CultureInfo.InvariantCulture);
            Assert.Equal(decimal.GetBits(expected), decimal.GetBits(actual));
        }

        [Theory]
        [InlineData("0")]
        [InlineData("-0")]
        [InlineData("-0.00")]
        [InlineData("42")]
        [InlineData("-42")]
        [InlineData("+42")]
        [InlineData("3.14")]
        [InlineData("1.50")]
        [InlineData("1.")]
        [InlineData(".5")]
        [InlineData("-.5")]
        [InlineData("00042")]
        [InlineData("0.000000000000000001")]
        [InlineData("0000000000000000001")]
        [InlineData("9999999999999999999")]
        [InlineData("1844674407.370955161")]
        [InlineData("12345678901234567.89")]
        public void AcceptsAndMatchesDecimalParseBitExactly(string text)
        {
            AssertBitEqualToDecimalParse(text);
        }

        [Theory]
        [InlineData("")]
        [InlineData("-")]
        [InlineData("+")]
        [InlineData(".")]
        [InlineData("-.")]
        [InlineData("1.2.3")]
        [InlineData("1,000")]
        [InlineData(" 1")]
        [InlineData("1 ")]
        [InlineData("1x")]
        [InlineData("--5")]
        [InlineData("1e5")]
        [InlineData("99999999999999999999")]
        [InlineData("0.0000000000000000001")]
        public void RejectsUnsupportedForms(string text)
        {
            Assert.False(FastDecimal.TryParse(Encoding.ASCII.GetBytes(text), out _), $"Expected FastDecimal to reject \"{text}\".");
        }

        [Fact]
        [SuppressMessage("Security", "CA5394:Do not use insecure randomness",
            Justification = "Needs a reproducible seeded PRNG for deterministic test content, not cryptographic randomness.")]
        public void MatchesDecimalParseOnRandomValues()
        {
            var rng = new Random(24680);
            StringBuilder text = new();
            for (int i = 0; i < 200_000; i++)
            {
                text.Clear();
                if (rng.Next(4) == 0)
                {
                    text.Append('-');
                }
                int integerDigits = rng.Next(1, 12);
                for (int d = integerDigits; d > 0; d--)
                {
                    text.Append((char)('0' + rng.Next(10)));
                }
                if (rng.Next(3) != 0)
                {
                    text.Append('.');
                    for (int d = rng.Next(19 - integerDigits); d > 0; d--)
                    {
                        text.Append((char)('0' + rng.Next(10)));
                    }
                }
                AssertBitEqualToDecimalParse(text.ToString());
            }
        }
    }
}
