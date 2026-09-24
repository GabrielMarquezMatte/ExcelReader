using System.Buffers.Text;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Tests.Writer
{
    public sealed class DoubleFormatTests
    {
        [Fact]
        [SuppressMessage("Security", "CA5394:Do not use insecure randomness",
            Justification = "Needs a reproducible seeded PRNG for deterministic test content, not cryptographic randomness.")]
        public void FormattedBytesMatchUtf8FormatterExactly()
        {
            Random rng = new(42);
            List<double> values = [0, -0.0, 1e9, 999_999_999.9999, -999_999_999.9999, 0.0001, 0.00001, 0.00015, 0.1 + 0.2, 1.0 / 3, 1e15, 2.5, -2.5, double.Epsilon, double.MaxValue];
            foreach (double edge in values.ToArray())
            {
                values.Add(Math.BitIncrement(edge));
                values.Add(Math.BitDecrement(edge));
            }
            for (int i = 0; i < 200_000; i++)
            {
                long digits = rng.NextInt64(-2_000_000_000_000, 2_000_000_000_000) / (long)Math.Pow(10, rng.Next(0, 12));
                values.Add(digits / Math.Pow(10, rng.Next(0, 7)));
                values.Add(rng.NextDouble() * Math.Pow(10, rng.Next(-8, 12)));
                values.Add(BitConverter.Int64BitsToDouble(rng.NextInt64()));
            }
            byte[] expected = new byte[64];
            byte[] actual = new byte[64];
            foreach (double value in values)
            {
                Utf8Formatter.TryFormat(value, expected, out int expectedLength);
                Assert.True(CellFormatter.TryFormatDouble(value, actual, out int actualLength));
                Assert.Equal(Encoding.ASCII.GetString(expected, 0, expectedLength), Encoding.ASCII.GetString(actual, 0, actualLength));
            }
        }

        [Fact]
        public void ReportsFailureWhenTheDestinationIsTooSmall()
        {
            Assert.False(CellFormatter.TryFormatDouble(-12.5, new byte[4], out _));
        }
    }
}
