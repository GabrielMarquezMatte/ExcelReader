using System.IO.Compression;
using ExcelReader.Core.Writer.Internal;
using ExcelReader.Core.Writer.Zip;

namespace ExcelReader.Tests.Writer.Zip
{
    public class ZipEntrySegmentTests
    {
        public static TheoryData<CompressionLevel, bool> Cases()
        {
            var data = new TheoryData<CompressionLevel, bool>();
            foreach (CompressionLevel level in Enum.GetValues<CompressionLevel>())
            {
                data.Add(level, true);
                data.Add(level, false);
            }
            return data;
        }

        private static Stream Target(MemoryStream ms, bool seekable)
        {
            if (seekable)
            {
                return ms;
            }
            return new ForwardOnlyWriteStream(ms);
        }

        private static (byte[] Segment, uint Crc) Segment(byte[] raw, CompressionLevel level)
        {
            using var buffer = new BiffBuffer(4096);
            SegmentCompressor.Compress(raw, level, buffer);
            return (buffer.Span.ToArray(), Crc32.Append(0, raw));
        }

        private static void AssertEntry(byte[] package, byte[] expected)
        {
            using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            ZipArchiveEntry entry = Assert.Single(zip.Entries);
            using Stream data = entry.Open();
            using var copy = new MemoryStream();
            data.CopyTo(copy);
            Assert.Equal(expected, copy.ToArray());
            Assert.Equal(Crc32Tests.BitwiseCrc(expected), entry.Crc32);
            Assert.Equal(expected.Length, entry.Length);
        }

        [Theory]
        [MemberData(nameof(Cases))]
        public void SerialWritesAndSegmentsInterleave(CompressionLevel level, bool seekable)
        {
            byte[] a = DeflateEncoderTests.SheetXml(1_500);
            byte[] b = DeflateEncoderTests.SheetXml(4_000);
            byte[] c = "<between/>"u8.ToArray();
            byte[] d = DeflateEncoderTests.SheetXml(2_500);
            byte[] e = DeflateEncoderTests.SheetXml(700);
            var ms = new MemoryStream();
            var zip = new ZipPackageWriter(Target(ms, seekable));
            using (ZipEntryStream entry = zip.OpenEntry("sheet.xml", level))
            {
                entry.Write(a);
                (byte[] segment, uint crc) = Segment(b, level);
                entry.AppendSegment(segment, crc, b.Length);
                entry.Write(c);
                (segment, crc) = Segment(d, level);
                entry.AppendSegment(segment, crc, d.Length);
                (segment, crc) = Segment([], level);
                entry.AppendSegment(segment, crc, 0);
                entry.Write(e);
            }
            zip.Finish();

            AssertEntry(ms.ToArray(), [.. a, .. b, .. c, .. d, .. e]);
        }

        [Theory]
        [MemberData(nameof(Cases))]
        public async Task ASegmentCanBeTheFirstAndOnlyContent(CompressionLevel level, bool seekable)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] raw = DeflateEncoderTests.SheetXml(3_000);
            var ms = new MemoryStream();
            var zip = new ZipPackageWriter(Target(ms, seekable));
            ZipEntryStream entry = zip.OpenEntry("sheet.xml", level);
            await using (entry)
            {
                (byte[] segment, uint crc) = Segment(raw, level);
                await entry.AppendSegmentAsync(segment, crc, raw.Length, ct);
            }
            await zip.FinishAsync(ct);

            AssertEntry(ms.ToArray(), raw);
        }

        // A zlib stream kept across a segment could point back at bytes the reader sees before the segment.
        [Fact]
        public void SerialWritesAfterASegmentStartAFreshStream()
        {
            byte[] repeated = DeflateEncoderTests.SheetXml(2_000);
            var ms = new MemoryStream();
            var zip = new ZipPackageWriter(ms);
            using (ZipEntryStream entry = zip.OpenEntry("sheet.xml", CompressionLevel.Optimal))
            {
                entry.Write(repeated);
                (byte[] segment, uint crc) = Segment("<other/>"u8.ToArray(), CompressionLevel.Optimal);
                entry.AppendSegment(segment, crc, 8);
                entry.Write(repeated);
            }
            zip.Finish();

            AssertEntry(ms.ToArray(), [.. repeated, .. "<other/>"u8, .. repeated]);
        }

        [Fact]
        public void AFailedAppendKillsTheEntryAndTheArchive()
        {
            byte[] raw = DeflateEncoderTests.SheetXml(20_000);
            var zip = new ZipPackageWriter(new FailingWriteStream(limit: 0));
            ZipEntryStream entry = zip.OpenEntry("sheet.xml", CompressionLevel.Fastest);
            (byte[] segment, uint crc) = Segment(raw, CompressionLevel.Fastest);

            Assert.Throws<IOException>(() => entry.AppendSegment(segment, crc, raw.Length));
            entry.Dispose();
            Assert.Throws<ObjectDisposedException>(() => entry.AppendSegment(segment, crc, raw.Length));
            Assert.Throws<InvalidOperationException>(zip.Finish);
            zip.Abandon();
        }

        [Fact]
        public async Task AFailedAsyncAppendKillsTheEntryAndTheArchive()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] raw = DeflateEncoderTests.SheetXml(20_000);
            var zip = new ZipPackageWriter(new FailingWriteStream(limit: 0));
            ZipEntryStream entry = zip.OpenEntry("sheet.xml", CompressionLevel.Fastest);
            (byte[] segment, uint crc) = Segment(raw, CompressionLevel.Fastest);

            await Assert.ThrowsAsync<IOException>(async () => await entry.AppendSegmentAsync(segment, crc, raw.Length, ct));
            await entry.DisposeAsync();
            await Assert.ThrowsAsync<ObjectDisposedException>(async () => await entry.AppendSegmentAsync(segment, crc, raw.Length, ct));
            await Assert.ThrowsAsync<InvalidOperationException>(async () => await zip.FinishAsync(ct));
            zip.Abandon();
        }
    }
}
