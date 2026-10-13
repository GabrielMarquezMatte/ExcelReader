using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using ExcelReader.Core.Reader.Zip.Inflate;
using ExcelReader.Core.Writer.Internal;
using ExcelReader.Core.Writer.Zip;

namespace ExcelReader.Tests.Writer.Zip
{
    public class SegmentCompressorTests
    {
        public static TheoryData<CompressionLevel> DeflateLevels()
        {
            return [CompressionLevel.Fastest, CompressionLevel.Optimal, CompressionLevel.SmallestSize];
        }

        private static byte[] Inflate(Func<Stream, Stream> open, byte[] compressed)
        {
            using Stream inflate = open(new MemoryStream(compressed));
            using var copy = new MemoryStream();
            inflate.CopyTo(copy);
            return copy.ToArray();
        }

        internal static void AssertInflatesTo(byte[] expected, byte[] compressed)
        {
            Assert.Equal(expected, Inflate(static s => new DeflateStream(s, CompressionMode.Decompress), compressed));
            Assert.Equal(expected, Inflate(static s => new InflateStream(s), compressed));
        }

        [Theory]
        [MemberData(nameof(DeflateLevels))]
        [SuppressMessage("Security", "CA5394:Do not use insecure randomness",
            Justification = "A seeded generator picks reproducible test inputs; nothing here is secret.")]
        public void SegmentsConcatenateIntoOneStream(CompressionLevel level)
        {
            var random = new Random(3);
            byte[] noise = new byte[70_000];
            random.NextBytes(noise);
            byte[][] parts =
            [
                DeflateEncoderTests.SheetXml(2_000),
                [],
                noise,
                [42],
                DeflateEncoderTests.SheetXml(9_000),
            ];
            using var output = new BiffBuffer(4096);
            foreach (byte[] part in parts)
            {
                SegmentCompressor.Compress(part, level, output);
            }
            output.Write(SegmentCompressor.FinalBlock);

            AssertInflatesTo([.. parts.SelectMany(static p => p)], output.Span.ToArray());
        }

        [Fact]
        public void NoCompressionCopiesTheBytes()
        {
            using var output = new BiffBuffer(64);
            SegmentCompressor.Compress("plain"u8, CompressionLevel.NoCompression, output);
            Assert.Equal("plain"u8.ToArray(), output.Span.ToArray());
        }

        [Fact]
        public void FlushSyncLeavesTheEncoderUsable()
        {
            byte[] first = DeflateEncoderTests.SheetXml(3_000);
            byte[] second = DeflateEncoderTests.SheetXml(1_000);
            using var encoder = new DeflateEncoder();
            using var output = new BiffBuffer(4096);
            encoder.Write(first, output);
            encoder.FlushSync(output);
            int aligned = output.Length;
            encoder.FlushSync(output);
            encoder.Write(second, output);
            encoder.Finish(output);

            Assert.Equal([0x00, 0x00, 0xFF, 0xFF], output.Span.Slice(aligned - 4, 4).ToArray());
            AssertInflatesTo([.. first, .. second], output.Span.ToArray());
        }

        [Fact]
        public void BufferSinkDropsWritesWhileDiscarding()
        {
            using var target = new BiffBuffer(64);
            using var sink = new BufferSink(target);
            sink.Write("kept"u8);
            sink.Discard = true;
            sink.Write("dropped"u8);
            Assert.Equal("kept"u8.ToArray(), target.Span.ToArray());
        }
    }
}
