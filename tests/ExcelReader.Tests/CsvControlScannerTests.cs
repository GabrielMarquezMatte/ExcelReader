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
            var scanner = new CsvControlScanner(delim, quote);
            scanner.Reset(data, len, start);
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


        [Fact]
        public void ContinuePreservesPendingMaskAcrossSimulatedRecordBoundary()
        {
            byte[] data = new byte[64];
            Array.Fill(data, (byte)'x');
            for (int i = 0; i < data.Length; i += 3)
            {
                data[i] = (byte)',';
            }
            List<int> expected = ExpectedStops(data, 0, data.Length, (byte)',', (byte)'"');

            var scanner = new CsvControlScanner((byte)',', (byte)'"');
            scanner.Reset(data, data.Length, 0);
            var actual = new List<int>();
            int stop = scanner.Next();
            int sinceBoundary = 0;
            while (stop >= 0)
            {
                actual.Add(stop);
                sinceBoundary++;
                if (sinceBoundary == expected.Count / 2)
                {
                    scanner.Continue(data, data.Length);
                }
                stop = scanner.Next();
            }
            Assert.Equal(expected, actual);
        }

        [Fact]
        public void SkipByteConsumesCrLfSecondByteWithoutReporting()
        {
            byte[] data = Encoding.UTF8.GetBytes("a,b\r\nc,d\n");
            List<int> stops = ScanWithCrLfSkip(data);
            Assert.Equal([1, 3, 6, 8], stops);
        }

        [Theory]
        [InlineData(15)]
        [InlineData(16)]
        [InlineData(17)]
        [InlineData(31)]
        [InlineData(32)]
        [InlineData(33)]
        public void SkipByteHandlesCrLfAtChunkBoundary(int offset)
        {
            byte[] data = new byte[offset + 10];
            Array.Fill(data, (byte)'x');
            data[offset] = (byte)'\r';
            data[offset + 1] = (byte)'\n';
            List<int> stops = ScanWithCrLfSkip(data);
            Assert.Equal([offset], stops);
        }

        private static List<int> ScanWithCrLfSkip(byte[] data)
        {
            var scanner = new CsvControlScanner((byte)',', (byte)'"');
            scanner.Reset(data, data.Length, 0);
            var stops = new List<int>();
            int stop = scanner.Next();
            while (stop >= 0)
            {
                stops.Add(stop);
                if (data[stop] == (byte)'\r' && stop + 1 < data.Length && data[stop + 1] == (byte)'\n')
                {
                    scanner.SkipByte(stop + 1);
                }
                stop = scanner.Next();
            }
            return stops;
        }

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
