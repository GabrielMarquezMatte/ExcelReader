using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using ExcelReader.Core.Reader.Zip.Inflate;
using static ExcelReader.Tests.Reader.Zip.InflateTestSupport;

namespace ExcelReader.Tests.Reader.Zip
{
    public class InflateStreamTests
    {
        [Fact]
        public void StoredBlockIsCopiedThrough()
        {
            byte[] stream = [0x01, 0x05, 0x00, 0xFA, 0xFF, (byte)'h', (byte)'e', (byte)'l', (byte)'l', (byte)'o'];

            Assert.Equal("hello"u8.ToArray(), Inflate(stream));
        }

        [Fact]
        public void EmptyFixedBlockProducesNothing()
        {
            Assert.Empty(Inflate([0x03, 0x00]));
        }

        [Fact]
        public void ZeroLengthStoredBlockIsSkipped()
        {
            byte[] stream = [0x00, 0x00, 0x00, 0xFF, 0xFF, 0x01, 0x02, 0x00, 0xFD, 0xFF, (byte)'o', (byte)'k'];

            Assert.Equal("ok"u8.ToArray(), Inflate(stream));
        }

        [Fact]
        public void BytesAfterTheFinalBlockAreIgnored()
        {
            byte[] raw = Generate(DataShape.SheetXml, 5000, seed: 1);
            byte[] stream = [.. Deflate(raw, CompressionLevel.Optimal), 0xDE, 0xAD, 0xBE, 0xEF];

            Assert.Equal(raw, Inflate(stream));
        }

        [Fact]
        public void ReservedBlockTypeIsRejected()
        {
            InvalidDataException ex = Assert.Throws<InvalidDataException>(() => Inflate([0x07, 0x00, 0x00, 0x00]));

            Assert.Contains("reserved block type", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void StoredBlockWithMismatchedLengthFieldsIsRejected()
        {
            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => Inflate([0x01, 0x05, 0x00, 0x00, 0x00, 1, 2, 3, 4, 5]));

            Assert.Contains("mismatched length", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void MatchReachingBeforeTheOutputStartIsRejected()
        {
            // Fixed block whose first symbol is a length-3 match at offset 1, with nothing before it.
            InvalidDataException ex = Assert.Throws<InvalidDataException>(() => Inflate([0x03, 0x02, 0x00, 0x00]));

            Assert.Contains("before the start", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void DynamicHeaderWhoseCodeLengthsOverrunIsRejected()
        {
            // Code-length code: symbols 0 and 18, one bit each. Two runs of 138 zeros exceed the 258 lengths declared.
            byte[] stream = DynamicHeader(precodeLens: [0, 0, 1, 1], codeLengthBits: [(1, 1), (127, 7), (1, 1), (127, 7)]);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(() => Inflate(stream));

            Assert.Contains("overrun", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void DynamicHeaderThatRepeatsBeforeAnyLengthIsRejected()
        {
            // Code-length code: symbols 0 and 16, one bit each. The first symbol is 16, "repeat the previous length".
            byte[] stream = DynamicHeader(precodeLens: [1, 0, 0, 1], codeLengthBits: [(1, 1), (0, 2)]);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(() => Inflate(stream));

            Assert.Contains("before defining one", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void DynamicHeaderWithOverSubscribedCodeLengthCodeIsRejected()
        {
            byte[] stream = DynamicHeader(precodeLens: [1, 1, 1, 1], codeLengthBits: []);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(() => Inflate(stream));

            Assert.Contains("code-length code", ex.Message, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(int.MaxValue)]
        [InlineData(1)]
        public void MatchAtTheFullWindowDistanceCopiesAcrossEverySlide(int sourceStep)
        {
            byte[] history = Generate(DataShape.Random, 32_768, seed: 3);
            byte[] expected = new byte[history.Length + (6000 * 258)];
            history.CopyTo(expected, 0);
            for (int i = history.Length; i < expected.Length; i++)
            {
                expected[i] = expected[i - 32_768];
            }

            Assert.Equal(expected, Inflate(StoredThenFarMatches(history, matches: 6000), sourceStep: sourceStep));
        }

        [Fact]
        public void MatchAtTheFullWindowDistanceWithOneByteTooFewIsRejected()
        {
            byte[] stream = StoredThenFarMatches(Generate(DataShape.Random, 32_767, seed: 3), matches: 1);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(() => Inflate(stream));

            Assert.Contains("before the start", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void EmptyInputIsRejected()
        {
            Assert.Throws<InvalidDataException>(() => Inflate([]));
        }

        [Theory]
        [InlineData(DataShape.SheetXml, CompressionLevel.Optimal)]
        [InlineData(DataShape.SheetXml, CompressionLevel.Fastest)]
        [InlineData(DataShape.Random, CompressionLevel.NoCompression)]
        [InlineData(DataShape.Zeros, CompressionLevel.Optimal)]
        [SuppressMessage("Security", "CA5394:Do not use insecure randomness", Justification = "Seeded cut positions, not security.")]
        public void StreamCutAnywhereIsRejected(DataShape shape, CompressionLevel level)
        {
            byte[] stream = Deflate(Generate(shape, 200_000, seed: 7), level);
            var random = new Random(11);
            int[] cuts = [1, 2, stream.Length / 2, stream.Length - 2, stream.Length - 1, .. Enumerable.Range(0, 40).Select(_ => random.Next(1, stream.Length))];

            foreach (int cut in cuts)
            {
                Assert.Throws<InvalidDataException>(() => Inflate(stream[..cut]));
            }
        }

        [Theory]
        [InlineData(DataShape.SheetXml, CompressionLevel.Optimal)]
        [InlineData(DataShape.SheetXml, CompressionLevel.Fastest)]
        [InlineData(DataShape.MixedRuns, CompressionLevel.Optimal)]
        public void StreamCutJustBeforeTheWindowFillsDeliversOnlyRealBytes(DataShape shape, CompressionLevel level)
        {
            // The window first fills at 98,304 bytes; cuts whose real output ends just short of it make
            // the zero padding decode across the boundary.
            byte[] raw = Generate(shape, 200_000, seed: 7);
            byte[] stream = Deflate(raw, level);
            int first = FirstCutReaching(stream, 98_304 - 300);
            int last = FirstCutReaching(stream, 98_304);

            for (int cut = first; cut <= last; cut++)
            {
                byte[] delivered = DeliveredBeforeRejection(stream[..cut]);
                Assert.True(raw.AsSpan().StartsWith(delivered), $"cut {cut} delivered bytes the stream never encoded");
            }
        }

        [Theory]
        [InlineData(1)]
        [InlineData(7)]
        [InlineData(4096)]
        [InlineData(1 << 20)]
        public void OutputIsTheSameForEveryReadSize(int readSize)
        {
            byte[] raw = Generate(DataShape.MixedRuns, 300_000, seed: 3);

            Assert.Equal(raw, Inflate(Deflate(raw, CompressionLevel.Optimal), readSize));
        }

        [Theory]
        [InlineData(1)]
        [InlineData(3)]
        [InlineData(997)]
        public void InputMayArriveInSmallPieces(int sourceStep)
        {
            byte[] raw = Generate(DataShape.SheetXml, 150_000, seed: 5);

            Assert.Equal(raw, Inflate(Deflate(raw, CompressionLevel.Optimal), readSize: 8192, sourceStep));
        }

        [Theory]
        [InlineData(DataShape.SheetXml, 1)]
        [InlineData(DataShape.RepeatedBlock, 997)]
        [InlineData(DataShape.Random, int.MaxValue)]
        public async Task ReadAsyncProducesTheSameBytesAsRead(DataShape shape, int sourceStep)
        {
            byte[] stream = Deflate(Generate(shape, 250_000, seed: 9), CompressionLevel.Optimal);

            byte[] sync = Inflate(stream, readSize: 5000, sourceStep);
            byte[] async = await InflateAsync(stream, readSize: 5000, sourceStep, TestContext.Current.CancellationToken);

            Assert.Equal(sync, async);
        }

        [Fact]
        public void StreamsDecodingInParallelDoNotInterfere()
        {
            byte[] raw = Generate(DataShape.SheetXml, 400_000, seed: 13);
            byte[] stream = Deflate(raw, CompressionLevel.Optimal);

            Parallel.For(0, 32, _ => Assert.True(raw.AsSpan().SequenceEqual(Inflate(stream, readSize: 3000))));
        }

        [Fact]
        public void DisposeDisposesTheInnerStream()
        {
            var inner = new SteppedStream([0x03, 0x00], int.MaxValue);

            new InflateStream(inner).Dispose();

            Assert.True(inner.Disposed);
        }

        [Fact]
        public void ReadAfterDisposeThrows()
        {
            var inflate = new InflateStream(new MemoryStream([0x03, 0x00]));
            inflate.Dispose();

            Assert.Throws<ObjectDisposedException>(() => inflate.Read(new byte[1]));
        }

        [Fact]
        public void DisposeTwiceIsHarmless()
        {
            var inflate = new InflateStream(new MemoryStream([0x03, 0x00]));

            inflate.Dispose();
            inflate.Dispose();

            Assert.False(inflate.CanRead);
        }

        // DeflateStream's output never shrinks as the cut grows, so a binary search finds the boundary.
        private static int FirstCutReaching(byte[] stream, int outputLength)
        {
            int low = 1;
            int high = stream.Length;
            while (low < high)
            {
                int mid = (low + high) / 2;
                using var reference = new DeflateStream(new MemoryStream(stream, 0, mid), CompressionMode.Decompress);
                using var output = new MemoryStream();
                reference.CopyTo(output);
                if (output.Length >= outputLength)
                {
                    high = mid;
                }
                else
                {
                    low = mid + 1;
                }
            }
            return low;
        }

        private static byte[] DeliveredBeforeRejection(byte[] stream)
        {
            using var inflate = new InflateStream(new MemoryStream(stream));
            using var output = new MemoryStream();
            byte[] buffer = new byte[4096];
            try
            {
                int read;
                while ((read = inflate.Read(buffer)) > 0)
                {
                    output.Write(buffer, 0, read);
                }
            }
            catch (InvalidDataException)
            {
                return output.ToArray();
            }
            throw new InvalidOperationException("A truncated stream was accepted.");
        }

        // A stored block holding the history, then a final fixed block of length-258 matches at offset
        // 32,768: symbol 285 and offset code 29 with all 13 extra bits set. Huffman codes go in bit-reversed.
        private static byte[] StoredThenFarMatches(byte[] history, int matches)
        {
            var bits = new BitWriter();
            bits.Write(1, 1);
            bits.Write(1, 2);
            for (int i = 0; i < matches; i++)
            {
                bits.Write(0b1010_0011, 8);
                bits.Write(0b1_0111, 5);
                bits.Write(8191, 13);
            }
            bits.Write(0, 7);
            int length = history.Length;
            return [0x00, (byte)length, (byte)(length >> 8), (byte)~length, (byte)(~length >> 8), .. history, .. bits.ToArray()];
        }

        // A final dynamic block declaring 257 literal/length codes and 1 offset code, with the given
        // lengths for the first four code-length-code symbols (16, 17, 18, 0) and raw bits after them.
        private static byte[] DynamicHeader(int[] precodeLens, (int Value, int Count)[] codeLengthBits)
        {
            var bits = new BitWriter();
            bits.Write(1, 1);
            bits.Write(2, 2);
            bits.Write(0, 5);
            bits.Write(0, 5);
            bits.Write(precodeLens.Length - 4, 4);
            foreach (int len in precodeLens)
            {
                bits.Write(len, 3);
            }
            foreach ((int value, int count) in codeLengthBits)
            {
                bits.Write(value, count);
            }
            return bits.ToArray();
        }

        private sealed class BitWriter
        {
            private readonly List<byte> _bytes = [];
            private int _used = 8;

            internal void Write(int value, int count)
            {
                for (int i = 0; i < count; i++)
                {
                    if (_used == 8)
                    {
                        _bytes.Add(0);
                        _used = 0;
                    }
                    _bytes[^1] |= (byte)(((value >> i) & 1) << _used);
                    _used++;
                }
            }

            internal byte[] ToArray()
            {
                return [.. _bytes];
            }
        }
    }
}
