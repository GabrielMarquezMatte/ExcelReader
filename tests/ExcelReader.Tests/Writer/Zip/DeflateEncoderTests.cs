using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Reader.Zip.Inflate;
using ExcelReader.Core.Writer.Internal;
using ExcelReader.Core.Writer.Zip;

namespace ExcelReader.Tests.Writer.Zip
{
    public class DeflateEncoderTests
    {
        internal static byte[] Encode(ReadOnlySpan<byte> input, Func<int, int> nextChunk)
        {
            using var encoder = new DeflateEncoder();
            using var output = new BiffBuffer(4096);
            while (!input.IsEmpty)
            {
                int take = Math.Clamp(nextChunk(input.Length), 1, input.Length);
                encoder.Write(input[..take], output);
                input = input[take..];
            }
            encoder.Finish(output);
            return output.Span.ToArray();
        }

        internal static byte[] SheetXml(int rows)
        {
            var builder = new StringBuilder();
            for (int row = 1; row <= rows; row++)
            {
                builder.Append(CultureInfo.InvariantCulture,
                    $"<row><c t=\"inlineStr\"><is><t>item {row % 97}</t></is></c><c><v>{row}</v></c><c><v>{row * 1.25}</v></c></row>");
            }
            return Encoding.UTF8.GetBytes(builder.ToString());
        }

        private static byte[] Inflate(byte[] compressed, bool library)
        {
            using var source = new MemoryStream(compressed);
            Stream inflate;
            if (library)
            {
                inflate = new InflateStream(source);
            }
            else
            {
                inflate = new DeflateStream(source, CompressionMode.Decompress);
            }
            using (inflate)
            {
                using var output = new MemoryStream();
                inflate.CopyTo(output);
                return output.ToArray();
            }
        }

        private static void AssertRoundTrips(byte[] input, byte[] compressed)
        {
            byte[] viaZlib = Inflate(compressed, library: false);
            byte[] viaLibrary = Inflate(compressed, library: true);
            Assert.True(viaZlib.AsSpan().SequenceEqual(input), $"DeflateStream gave {viaZlib.Length} bytes for {input.Length}");
            Assert.True(viaLibrary.AsSpan().SequenceEqual(input), $"InflateStream gave {viaLibrary.Length} bytes for {input.Length}");
        }

        [SuppressMessage("Security", "CA5394:Do not use insecure randomness",
            Justification = "A seeded generator picks reproducible test inputs; nothing here is secret.")]
        private static byte[] RandomBytes(int length, int seed)
        {
            byte[] bytes = new byte[length];
            new Random(seed).NextBytes(bytes);
            return bytes;
        }

        [Fact]
        public void EmptyInputIsOneEmptyFinalBlock()
        {
            byte[] compressed = Encode([], static _ => 1);
            Assert.Equal([0x03, 0x00], compressed);
            AssertRoundTrips([], compressed);
        }

        [Fact]
        public void SingleByteRoundTrips()
        {
            byte[] input = [(byte)'x'];
            AssertRoundTrips(input, Encode(input, static _ => 1));
        }

        [Fact]
        public void SheetXmlRoundTripsAndShrinks()
        {
            byte[] input = SheetXml(50_000);
            byte[] compressed = Encode(input, static _ => 64 * 1024);
            AssertRoundTrips(input, compressed);
            Assert.True(compressed.Length < input.Length / 4, $"{compressed.Length} of {input.Length}");
        }

        [Fact]
        public void TenMegabytesOfZerosUseLongMatches()
        {
            byte[] input = new byte[10 << 20];
            byte[] compressed = Encode(input, static _ => 1 << 20);
            AssertRoundTrips(input, compressed);
            Assert.True(compressed.Length < input.Length / 50, $"{compressed.Length} of {input.Length}");
        }

        [Fact]
        public void RandomBytesFallBackToStoredBlocks()
        {
            byte[] input = RandomBytes(1 << 20, seed: 3);
            byte[] compressed = Encode(input, static _ => 64 * 1024);
            AssertRoundTrips(input, compressed);
            long bound = input.Length + (5L * ((input.Length / DeflateEncoder.MaxBlockLength) + 1)) + 16;
            Assert.True(compressed.Length <= bound, $"{compressed.Length} > {bound}");
        }

        [Fact]
        public void CjkTextRoundTrips()
        {
            var builder = new StringBuilder();
            for (int i = 0; i < 20_000; i++)
            {
                builder.Append(CultureInfo.InvariantCulture, $"日本語のテキスト{i % 313}、中文文本。");
            }
            byte[] input = Encoding.UTF8.GetBytes(builder.ToString());
            byte[] compressed = Encode(input, static _ => 4096);
            AssertRoundTrips(input, compressed);
            Assert.True(compressed.Length < input.Length, $"{compressed.Length} of {input.Length}");
        }

        [Fact]
        [SuppressMessage("Security", "CA5394:Do not use insecure randomness",
            Justification = "A seeded generator picks reproducible test inputs; nothing here is secret.")]
        public void RandomWriteSizesRoundTrip()
        {
            byte[] input = [.. SheetXml(30_000), .. RandomBytes(300_000, seed: 4), .. SheetXml(30_000)];
            var random = new Random(9);
            AssertRoundTrips(input, Encode(input, _ => random.Next(1, 1 << 20)));
        }

        [Fact]
        public void OneHugeWriteSlidesTheWindowMidWrite()
        {
            byte[] input = SheetXml(120_000);
            Assert.True(input.Length > 4 << 20);
            AssertRoundTrips(input, Encode(input, static remaining => remaining));
        }

        [Fact]
        public void ByteAtATimeWritesRoundTrip()
        {
            byte[] input = SheetXml(3_000);
            AssertRoundTrips(input, Encode(input, static _ => 1));
        }

        [Fact]
        public void MatchesAtAndPastTheWindowEdgeRoundTrip()
        {
            byte[] history = RandomBytes(32_768, seed: 6);
            byte[] input = [.. history, .. history.AsSpan(0, 1_000), .. history];
            AssertRoundTrips(input, Encode(input, static _ => 64 * 1024));
        }

        [Fact]
        public void FarMatchAtBlockStartSurvivesPendingBits()
        {
            byte[] run = RandomBytes(300, seed: 1);
            run[0] = 0xAB;
            for (int prefix = 0; prefix < 256; prefix++)
            {
                foreach (int distance in new[] { 30_000, 22_000, 18_000 })
                {
                    byte[] input = new byte[DeflateEncoder.MaxBlockLength + 2_000];
                    RandomBytes(prefix, seed: prefix).CopyTo(input, 0);
                    run.CopyTo(input, DeflateEncoder.MaxBlockLength - distance);
                    run.AsSpan(0, 200).CopyTo(input.AsSpan(DeflateEncoder.MaxBlockLength));
                    RandomBytes(1_800, seed: 1_000 + prefix).CopyTo(input, DeflateEncoder.MaxBlockLength + 200);
                    input[DeflateEncoder.MaxBlockLength + 200] = (byte)(run[200] ^ 0xFF);
                    AssertRoundTrips(input, Encode(input, static remaining => remaining));
                }
            }
        }

        [Fact]
        public void FixedCodesFollowRfc1951()
        {
            Assert.Equal(0x188u, FixedHuffmanCodes.Literals[0]);
            Assert.Equal(0x269u, FixedHuffmanCodes.Literals[144]);
            Assert.Equal(0x807u, FixedHuffmanCodes.Lengths[3]);
            Assert.Equal(0x1468u, FixedHuffmanCodes.Lengths[258]);
            Assert.Equal(0x5u, FixedHuffmanCodes.Distances[1]);
            Assert.Equal(0x7FFEF2u, FixedHuffmanCodes.Distances[32_768]);
        }
    }
}
