using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests.Reader
{
    public class CellStoredDecimalTests
    {
        [Theory]
        [InlineData(1.23, "1.23")]
        [InlineData(0.1 + 0.2, "0.30000000000000004")]
        [InlineData(1e-7, "0.0000001")]
        [InlineData(-45.5, "-45.5")]
        [InlineData(123456789.0, "123456789")]
        [InlineData(1e20, "100000000000000000000")]
        [InlineData(-0.07, "-0.07")]
        [InlineData(123456789012.345, "123456789012.345")]
        [InlineData(1.5e-20, "0.000000000000000000015")]
        public void StoredNumberParsesAsTheDecimalOfItsXlsxText(double number, string expected)
        {
            string text = number.ToString("R", CultureInfo.InvariantCulture);
            var stored = new Cell(CellType.Number, Encoding.ASCII.GetBytes(text), number, hasNumber: true, styleIndex: 0);
            var textBacked = new Cell(CellType.Number, Encoding.ASCII.GetBytes(text));

            Assert.True(stored.TryParse(CultureInfo.InvariantCulture, out decimal fromStored));
            Assert.True(textBacked.TryParse(CultureInfo.InvariantCulture, out decimal fromText));

            Assert.Equal(decimal.Parse(expected, CultureInfo.InvariantCulture), fromStored);
            Assert.Equal(fromText, fromStored);
        }

        [Fact]
        [SuppressMessage("Security", "CA5394:Do not use insecure randomness",
            Justification = "Needs a reproducible seeded PRNG for deterministic test content, not cryptographic randomness.")]
        public void RandomStoredNumbersMatchTheirShortestText()
        {
            var random = new Random(20261022);
            for (int i = 0; i < 20_000; i++)
            {
                double number = i % 2 == 0
                    ? Math.Round((random.NextDouble() - 0.5) * Math.Pow(10, random.Next(-3, 12)), random.Next(0, 8))
                    : BitConverter.Int64BitsToDouble(random.NextInt64()) % 1e18;
                if (!double.IsFinite(number))
                {
                    continue;
                }
                string text = number.ToString("R", CultureInfo.InvariantCulture);
                var stored = new Cell(CellType.Number, ReadOnlySpan<byte>.Empty, number, hasNumber: true, styleIndex: 0);

                Assert.True(stored.TryParse(CultureInfo.InvariantCulture, out decimal fromStored), text);
                Assert.Equal(decimal.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture), fromStored);
            }
        }

        [Theory]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        [InlineData(1e30)]
        public void StoredNumberOutsideDecimalRangeDoesNotParse(double number)
        {
            var stored = new Cell(CellType.Number, ReadOnlySpan<byte>.Empty, number, hasNumber: true, styleIndex: 0);

            Assert.False(stored.TryParse(CultureInfo.InvariantCulture, out decimal _));
        }
    }
}
