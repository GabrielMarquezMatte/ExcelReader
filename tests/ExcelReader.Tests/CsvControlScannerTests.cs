using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class CsvControlScannerTests
    {
        private static List<int> ExpectedStops(byte[] data, int start, int len, byte delim, byte quote)
        {
            var stops = new List<int>();
            for (int i = start; i < len; i++)
            {
                byte b = data[i];
                if (b == delim || b == quote || b == (byte)'\r' || b == (byte)'\n')
                {
                    stops.Add(i);
                }
            }
            return stops;
        }

        private static List<int> ActualStops(byte[] data, int start, int len, byte delim, byte quote)
        {
            var stops = new List<int>();
            var scanner = new CsvControlScanner(data.AsSpan(), start, len, delim, quote);
            int stop = scanner.Next();
            while (stop >= 0)
            {
                stops.Add(stop);
                stop = scanner.Next();
            }
            return stops;
        }

        [Fact]
        public void EmptyRangeYieldsNoStops()
        {
            byte[] data = "abc"u8.ToArray();
            Assert.Empty(ActualStops(data, 3, 3, (byte)',', (byte)'"'));
        }

        [Fact]
        public void NoControlBytesYieldsNoStops()
        {
            byte[] data = "abcdefghijklmnopqrstuvwxyz0123456789"u8.ToArray();
            Assert.Empty(ActualStops(data, 0, data.Length, (byte)',', (byte)'"'));
        }

        [Fact]
        public void ShortInputUsesScalarTail()
        {
            byte[] data = "a,b,c\n"u8.ToArray();
            List<int> expected = ExpectedStops(data, 0, data.Length, (byte)',', (byte)'"');
            List<int> actual = ActualStops(data, 0, data.Length, (byte)',', (byte)'"');
            Assert.Equal(expected, actual);
        }

        [Fact]
        public void StopsInsideOneChunkAreReportedInOrder()
        {
            byte[] data = Encoding.UTF8.GetBytes("aa,bb,cc,dd,ee,ff\n");
            List<int> expected = ExpectedStops(data, 0, data.Length, (byte)',', (byte)'"');
            List<int> actual = ActualStops(data, 0, data.Length, (byte)',', (byte)'"');
            Assert.Equal(expected, actual);
        }

        [Theory]
        [InlineData(15)]
        [InlineData(16)]
        [InlineData(17)]
        [InlineData(31)]
        [InlineData(32)]
        [InlineData(33)]
        public void StopAtChunkBoundaryIsReported(int offset)
        {
            byte[] data = new byte[offset + 10];
            Array.Fill(data, (byte)'x');
            data[offset] = (byte)',';
            List<int> expected = ExpectedStops(data, 0, data.Length, (byte)',', (byte)'"');
            List<int> actual = ActualStops(data, 0, data.Length, (byte)',', (byte)'"');
            Assert.Equal(expected, actual);
        }

        [Fact]
        public void StopsSpanningManyChunksAreReportedInOrder()
        {
            byte[] data = new byte[200];
            Array.Fill(data, (byte)'x');
            for (int i = 0; i < data.Length; i += 7)
            {
                data[i] = (byte)',';
            }
            List<int> expected = ExpectedStops(data, 0, data.Length, (byte)',', (byte)'"');
            List<int> actual = ActualStops(data, 0, data.Length, (byte)',', (byte)'"');
            Assert.Equal(expected, actual);
        }

        [Fact]
        public void ControlByteInTailAfterFullChunksIsReported()
        {
            byte[] data = new byte[70];
            Array.Fill(data, (byte)'x');
            data[69] = (byte)'\n';
            List<int> expected = ExpectedStops(data, 0, data.Length, (byte)',', (byte)'"');
            List<int> actual = ActualStops(data, 0, data.Length, (byte)',', (byte)'"');
            Assert.Equal(expected, actual);
        }

        [Fact]
        public void StartOffsetIsHonored()
        {
            byte[] data = Encoding.UTF8.GetBytes("first,line\nsecond,line\n");
            int start = data.AsSpan().IndexOf((byte)'\n') + 1;
            List<int> expected = ExpectedStops(data, start, data.Length, (byte)',', (byte)'"');
            List<int> actual = ActualStops(data, start, data.Length, (byte)',', (byte)'"');
            Assert.Equal(expected, actual);
            Assert.DoesNotContain(expected, i => i < start);
        }

        [Fact]
        public void LenShorterThanBufferIsRespected()
        {
            byte[] data = new byte[128];
            Array.Fill(data, (byte)'x');
            data[80] = (byte)',';
            data[100] = (byte)',';
            int len = 70;
            List<int> actual = ActualStops(data, 0, len, (byte)',', (byte)'"');
            Assert.Empty(actual);
        }

        [Theory]
        [InlineData((byte)';', (byte)'\'')]
        [InlineData((byte)'\t', (byte)'"')]
        [InlineData((byte)'|', (byte)'`')]
        public void CustomDelimiterAndQuoteAreMatched(byte delim, byte quote)
        {
            byte[] data = [(byte)'a', delim, quote, (byte)'b', (byte)'\r', (byte)'\n'];
            List<int> expected = ExpectedStops(data, 0, data.Length, delim, quote);
            List<int> actual = ActualStops(data, 0, data.Length, delim, quote);
            Assert.Equal(expected, actual);
        }

        [Fact]
        public void MultibyteUtf8IsNotMistakenForControl()
        {
            byte[] data = Encoding.UTF8.GetBytes("São Paulo,Ribeirão\n");
            List<int> expected = ExpectedStops(data, 0, data.Length, (byte)',', (byte)'"');
            List<int> actual = ActualStops(data, 0, data.Length, (byte)',', (byte)'"');
            Assert.Equal(expected, actual);
        }

        // Deliberately not cryptographically secure — CA5394 doesn't apply: a seeded, reproducible
        // PRNG is exactly what makes a parity failure reproducible, unlike a CSPRNG would be.
        [SuppressMessage("Security", "CA5394:Do not use insecure randomness",
            Justification = "Randomized parity test needs a reproducible seeded PRNG, not cryptographic randomness.")]
        [Theory]
        [InlineData(1)]
        [InlineData(2)]
        [InlineData(3)]
        [InlineData(4)]
        [InlineData(5)]
        [InlineData(6)]
        [InlineData(7)]
        [InlineData(8)]
        public void RandomizedParityAgainstScalarReference(int seed)
        {
            var random = new Random(seed);
            byte[] data = new byte[4096];
            random.NextBytes(data);
            // Push control-byte density up: replace ~1 in 6 bytes with one of the four control bytes.
            byte[] controls = [(byte)',', (byte)'"', (byte)'\r', (byte)'\n'];
            foreach (ref var b in data.AsSpan())
            {
                if (random.Next(6) == 0)
                {
                    b = controls[random.Next(controls.Length)];
                }
            }
            List<int> expected = ExpectedStops(data, 0, data.Length, (byte)',', (byte)'"');
            List<int> actual = ActualStops(data, 0, data.Length, (byte)',', (byte)'"');
            Assert.Equal(expected, actual);
        }
    }
}
