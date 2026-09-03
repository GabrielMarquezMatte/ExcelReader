using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Tests
{
    public sealed class OleCompoundWriterTests
    {
        private static CfbStreamSpec Spec(string name, byte[] content)
        {
            return new CfbStreamSpec(
                name,
                content.Length,
                stream => stream.Write(content),
                async (stream, ct) => await stream.WriteAsync(content, ct).ConfigureAwait(false));
        }

        private static byte[] Pattern(int length, byte seed)
        {
            byte[] content = new byte[length];
            for (int i = 0; i < length; i++)
            {
                content[i] = (byte)(seed + i);
            }
            return content;
        }

        private static byte[] ReadBack(byte[] container, string name, out long declaredSize)
        {
            using var source = new MemoryStream(container, writable: false);
            using CfbContainer cfb = CfbContainer.Parse(source, ownsSource: false, ExcelReaderOptions.Default);
            declaredSize = cfb.StreamLength(name);
            return cfb.ReadStream(name, 64 * 1024 * 1024);
        }

        [Fact]
        public void Write_MiniAndBigStream_BothReadBackByteIdentical()
        {
            byte[] small = Pattern(1300, 0x11);   // below the 4096 mini cutoff
            byte[] big = Pattern(9000, 0x77);     // above it

            using var ms = new MemoryStream();
            OleCompoundWriter.Write(ms, [Spec("EncryptionInfo", small), Spec("EncryptedPackage", big)]);

            Assert.Equal(small, ReadBack(ms.ToArray(), "EncryptionInfo", out long smallSize));
            Assert.Equal(1300, smallSize);
            Assert.Equal(big, ReadBack(ms.ToArray(), "EncryptedPackage", out long bigSize));
            Assert.Equal(9000, bigSize);
        }

        [Theory]
        [InlineData(1)]
        [InlineData(63)]
        [InlineData(64)]
        [InlineData(65)]
        [InlineData(4095)]
        [InlineData(4096)]
        [InlineData(4097)]
        public void Write_StreamSizeAcrossSectorBoundaries_ReadsBackByteIdentical(int size)
        {
            byte[] content = Pattern(size, 0x2A);

            using var ms = new MemoryStream();
            OleCompoundWriter.Write(ms, [Spec("OnlyStream", content)]);

            Assert.Equal(content, ReadBack(ms.ToArray(), "OnlyStream", out long declared));
            Assert.Equal(size, declared);
        }

        [Fact]
        public void Write_ManySmallStreams_EachChainsToItsOwnMiniSectors()
        {
            CfbStreamSpec[] specs = new CfbStreamSpec[20];
            byte[][] contents = new byte[20][];
            for (int i = 0; i < specs.Length; i++)
            {
                contents[i] = Pattern(100 + (i * 37), (byte)(i + 1));
                specs[i] = Spec($"Stream{i:D2}", contents[i]);
            }

            using var ms = new MemoryStream();
            OleCompoundWriter.Write(ms, specs);

            byte[] container = ms.ToArray();
            for (int i = 0; i < specs.Length; i++)
            {
                Assert.Equal(contents[i], ReadBack(container, $"Stream{i:D2}", out _));
            }
        }

        [Fact]
        public async Task WriteAsync_MatchesTheSynchronousWriterByteForByte()
        {
            byte[] small = Pattern(1300, 0x11);
            byte[] big = Pattern(9000, 0x77);

            using var syncMs = new MemoryStream();
            OleCompoundWriter.Write(syncMs, [Spec("EncryptionInfo", small), Spec("EncryptedPackage", big)]);

            using var asyncMs = new MemoryStream();
            await OleCompoundWriter.WriteAsync(asyncMs, [Spec("EncryptionInfo", small), Spec("EncryptedPackage", big)],
                TestContext.Current.CancellationToken);

            Assert.Equal(syncMs.ToArray(), asyncMs.ToArray());
        }

        // A stream past 109 FAT sectors (109 * 128 * 512 ≈ 7.1 MB) is the only way to reach the DIFAT
        // branch, so the size here is load-bearing, not arbitrary.
        [Fact]
        public void Write_StreamLargeEnoughToNeedDifat_ReadsBackByteIdentical()
        {
            byte[] content = Pattern(7_400_000, 0x5C);

            using var ms = new MemoryStream();
            OleCompoundWriter.Write(ms, [Spec("Big", content)]);

            Assert.Equal(content, ReadBack(ms.ToArray(), "Big", out _));
        }

        [Fact]
        public void Write_EmptyStream_Throws()
        {
            using var ms = new MemoryStream();
            Assert.Throws<ArgumentException>(() => OleCompoundWriter.Write(ms, [Spec("Empty", [])]));
        }
    }
}
