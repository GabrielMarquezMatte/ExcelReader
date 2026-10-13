using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Writer.Zip;

namespace ExcelReader.Tests.Writer.Zip
{
    public class ZipPackageWriterTests
    {
        private const int ChunkSize = 8192;
        private static readonly (string Name, byte[] Content)[] Entries = BuildEntries();

        [SuppressMessage("Security", "CA5394:Do not use insecure randomness",
            Justification = "A seeded generator picks reproducible test inputs; nothing here is secret.")]
        private static (string Name, byte[] Content)[] BuildEntries()
        {
            var random = new Random(11);
            var digits = new StringBuilder();
            while (digits.Length < 1_000_000)
            {
                digits.Append(random.Next().ToString(CultureInfo.InvariantCulture)).Append(',');
            }
            return
            [
                ("a.xml", "<a>hello hello hello</a>"u8.ToArray()),
                ("empty.xml", []),
                ("big.xml", Encoding.ASCII.GetBytes(digits.ToString())),
                ("ünïcode.xml", "<u/>"u8.ToArray()),
            ];
        }

        public static TheoryData<CompressionLevel, bool, int, bool> LayoutCases()
        {
            var data = new TheoryData<CompressionLevel, bool, int, bool>();
            foreach (CompressionLevel level in Enum.GetValues<CompressionLevel>())
            {
                data.Add(level, true, 0, false);
                data.Add(level, true, 0, true);
                data.Add(level, false, 0, false);
                data.Add(level, false, 0, true);
                data.Add(level, true, 100, false);
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

        private static void WriteInChunks(Stream entry, byte[] content)
        {
            for (int offset = 0; offset < content.Length; offset += ChunkSize)
            {
                entry.Write(content, offset, Math.Min(ChunkSize, content.Length - offset));
            }
        }

        private static byte[] WriteWithZipArchive(CompressionLevel level, bool seekable, int prefix)
        {
            var ms = new MemoryStream();
            ms.Write(new byte[prefix]);
            using (var zip = new ZipArchive(Target(ms, seekable), ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach ((string name, byte[] content) in Entries)
                {
                    using Stream entry = zip.CreateEntry(name, level).Open();
                    WriteInChunks(entry, content);
                }
            }
            return ms.ToArray();
        }

        private static async Task<byte[]> WriteWithPackageAsync(CompressionLevel level, bool seekable, int prefix, bool useAsync)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var ms = new MemoryStream();
            ms.Write(new byte[prefix]);
            var zip = new ZipPackageWriter(Target(ms, seekable));
            foreach ((string name, byte[] content) in Entries)
            {
                Stream entry = zip.OpenEntry(name, level);
                if (!useAsync)
                {
                    WriteInChunks(entry, content);
                    entry.Dispose();
                    continue;
                }
                for (int offset = 0; offset < content.Length; offset += ChunkSize)
                {
                    await entry.WriteAsync(content.AsMemory(offset, Math.Min(ChunkSize, content.Length - offset)), ct);
                }
                await entry.DisposeAsync();
            }
            if (useAsync)
            {
                await zip.FinishAsync(ct);
            }
            else
            {
                zip.Finish();
            }
            return ms.ToArray();
        }

        private static void AssertContentAndCrc(byte[] zipBytes)
        {
            using var zip = new ZipArchive(new MemoryStream(zipBytes), ZipArchiveMode.Read);
            Assert.Equal(Entries.Length, zip.Entries.Count);
            foreach ((string name, byte[] content) in Entries)
            {
                ZipArchiveEntry entry = zip.GetEntry(name)!;
                using Stream data = entry.Open();
                using var copy = new MemoryStream();
                data.CopyTo(copy);
                Assert.True(copy.ToArray().AsSpan().SequenceEqual(content), name);
                Assert.Equal(Crc32Tests.BitwiseCrc(content), entry.Crc32);
            }
        }

        private static void AssertSameShape(ZipCentral expected, ZipCentral actual, byte[] expectedZip, byte[] actualZip)
        {
            Assert.Equal(expected.Name, actual.Name);
            Assert.Equal(expected.MadeBy, actual.MadeBy);
            Assert.Equal(expected.Version, actual.Version);
            Assert.Equal(expected.Flags, actual.Flags);
            Assert.Equal(expected.Method, actual.Method);
            Assert.Equal(expected.Crc, actual.Crc);
            Assert.Equal(expected.UncompressedSize, actual.UncompressedSize);
            Assert.Equal(expected.Extra, actual.Extra);
            ZipLocal expectedLocal = ZipLayout.ReadLocal(expectedZip, expected.LocalOffset);
            ZipLocal actualLocal = ZipLayout.ReadLocal(actualZip, actual.LocalOffset);
            Assert.Equal(expectedLocal.Version, actualLocal.Version);
            Assert.Equal(expectedLocal.Flags, actualLocal.Flags);
            Assert.Equal(expectedLocal.Method, actualLocal.Method);
            Assert.Equal(expectedLocal.Crc, actualLocal.Crc);
            Assert.Equal(expectedLocal.UncompressedSize, actualLocal.UncompressedSize);
            Assert.Equal(expectedLocal.CompressedSize == 0, actualLocal.CompressedSize == 0);
            if (actualLocal.CompressedSize != 0)
            {
                Assert.Equal(actual.CompressedSize, actualLocal.CompressedSize);
            }
        }

        [Theory]
        [MemberData(nameof(LayoutCases))]
        public async Task LayoutMatchesZipArchive(CompressionLevel level, bool seekable, int prefix, bool useAsync)
        {
            byte[] expected = WriteWithZipArchive(level, seekable, prefix);
            byte[] actual = await WriteWithPackageAsync(level, seekable, prefix, useAsync);
            List<ZipCentral> expectedEntries = ZipLayout.ReadCentral(expected);
            List<ZipCentral> actualEntries = ZipLayout.ReadCentral(actual);
            Assert.Equal(expectedEntries.Count, actualEntries.Count);
            for (int i = 0; i < expectedEntries.Count; i++)
            {
                AssertSameShape(expectedEntries[i], actualEntries[i], expected, actual);
            }
            if (level == CompressionLevel.NoCompression)
            {
                Assert.Equal(ZipLayout.WithoutTimestamps(expected), ZipLayout.WithoutTimestamps(actual));
            }
            AssertContentAndCrc(actual);
        }

        [Fact]
        public void ZeroEntriesWriteTheEmptyArchive()
        {
            var expected = new MemoryStream();
            using (new ZipArchive(expected, ZipArchiveMode.Create, leaveOpen: true))
            {
            }
            var actual = new MemoryStream();
            new ZipPackageWriter(actual).Finish();
            Assert.Equal(expected.ToArray(), actual.ToArray());
            using var read = new ZipArchive(new MemoryStream(actual.ToArray()), ZipArchiveMode.Read);
            Assert.Empty(read.Entries);
        }

        [Fact]
        public void OpeningASecondEntryWhileOneIsOpenThrows()
        {
            var zip = new ZipPackageWriter(new MemoryStream());
            using Stream first = zip.OpenEntry("a.xml", CompressionLevel.Fastest);
            Assert.Throws<InvalidOperationException>(() => zip.OpenEntry("b.xml", CompressionLevel.Fastest));
            Assert.Throws<InvalidOperationException>(zip.Finish);
        }

        [Fact]
        public void UndefinedCompressionLevelIsRejected()
        {
            var zip = new ZipPackageWriter(new MemoryStream());
            Assert.Throws<ArgumentOutOfRangeException>(() => zip.OpenEntry("a.xml", (CompressionLevel)42));
        }

        [Fact]
        public async Task DisposingAnEntryAgainIsHarmless()
        {
            var ms = new MemoryStream();
            var zip = new ZipPackageWriter(ms);
            Stream entry = zip.OpenEntry("a.xml", CompressionLevel.Fastest);
            entry.Write("<a/>"u8);
            await entry.DisposeAsync();
            entry.Dispose();
            await entry.DisposeAsync();
            Assert.Throws<ObjectDisposedException>(() => entry.Write("x"u8));
            zip.Finish();
            Assert.Single(ZipLayout.ReadCentral(ms.ToArray()));
        }

        [Fact]
        public void WriteByteBuildsTheSameEntry()
        {
            byte[] content = DeflateEncoderTests.SheetXml(200);
            var ms = new MemoryStream();
            var zip = new ZipPackageWriter(ms);
            using (Stream entry = zip.OpenEntry("bytes.xml", CompressionLevel.Fastest))
            {
                foreach (byte value in content)
                {
                    entry.WriteByte(value);
                }
            }
            zip.Finish();
            using var read = new ZipArchive(new MemoryStream(ms.ToArray()), ZipArchiveMode.Read);
            ZipArchiveEntry stored = read.GetEntry("bytes.xml")!;
            using var copy = new MemoryStream();
            using (Stream data = stored.Open())
            {
                data.CopyTo(copy);
            }
            Assert.True(copy.ToArray().AsSpan().SequenceEqual(content));
            Assert.Equal(Crc32Tests.BitwiseCrc(content), stored.Crc32);
        }

        [Fact]
        public void AFailedWriteLeavesNoCentralDirectory()
        {
            var output = new FailingWriteStream(limit: 100_000);
            var zip = new ZipPackageWriter(output);
            Stream entry = zip.OpenEntry("big.xml", CompressionLevel.NoCompression);
            byte[] chunk = new byte[64 * 1024];
            Assert.Throws<IOException>(() =>
            {
                for (int i = 0; i < 10; i++)
                {
                    entry.Write(chunk);
                }
            });
            Assert.Throws<IOException>(entry.Dispose);
            Assert.Throws<InvalidOperationException>(zip.Finish);
            zip.Abandon();
            zip.Abandon();
            Assert.Equal(-1, output.Written.ToArray().AsSpan().IndexOf("PK\u0001\u0002"u8));
            Assert.Throws<ObjectDisposedException>(() => zip.OpenEntry("next.xml", CompressionLevel.Fastest));
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void AnEntryOver4GiBUsesZip64AndSoDoLaterOffsets(bool seekable)
        {
            const long Size = (1L << 32) + (1 << 20);
            var sink = new CapturingStream(seekable);
            var zip = new ZipPackageWriter(sink);
            byte[] chunk = new byte[1 << 20];
            uint crc = 0;
            using (Stream big = zip.OpenEntry("big.bin", CompressionLevel.NoCompression))
            {
                for (long written = 0; written < Size; written += chunk.Length)
                {
                    big.Write(chunk);
                    crc = Crc32.Append(crc, chunk);
                }
            }
            using (Stream after = zip.OpenEntry("after.xml", CompressionLevel.Fastest))
            {
                after.Write("<after/>"u8);
            }
            zip.Finish();

            ZipLocal local = ZipLayout.ReadLocal(sink.Head, 0);
            Assert.Equal(ZipRecords.FlagDataDescriptor, local.Flags & ZipRecords.FlagDataDescriptor);
            Assert.Equal(seekable ? 45 : 20, local.Version);

            byte[] tail = sink.Tail;
            long descriptorAt = ZipRecords.LocalHeaderLength + "big.bin".Length + Size;
            long d = descriptorAt - CapturingStream.TailStart;
            Assert.Equal(0x0807_4B50u, ZipLayout.U32(tail, d));
            Assert.Equal(crc, ZipLayout.U32(tail, d + 4));
            Assert.Equal((ulong)Size, ZipLayout.U64(tail, d + 8));
            Assert.Equal((ulong)Size, ZipLayout.U64(tail, d + 16));

            long end = tail.Length - 22;
            Assert.Equal(0x0605_4B50u, ZipLayout.U32(tail, end));
            Assert.Equal(uint.MaxValue, ZipLayout.U32(tail, end + 16));
            long locator = end - 20;
            Assert.Equal(0x0706_4B50u, ZipLayout.U32(tail, locator));
            long zip64End = (long)ZipLayout.U64(tail, locator + 8) - CapturingStream.TailStart;
            Assert.Equal(0x0606_4B50u, ZipLayout.U32(tail, zip64End));
            Assert.Equal(2UL, ZipLayout.U64(tail, zip64End + 32));
            long central = (long)ZipLayout.U64(tail, zip64End + 48) - CapturingStream.TailStart;

            Assert.Equal(0x0201_4B50u, ZipLayout.U32(tail, central));
            Assert.Equal(45, ZipLayout.U16(tail, central + 6));
            Assert.Equal(uint.MaxValue, ZipLayout.U32(tail, central + 20));
            Assert.Equal(uint.MaxValue, ZipLayout.U32(tail, central + 24));
            Assert.Equal(20, ZipLayout.U16(tail, central + 30));
            Assert.Equal(0u, ZipLayout.U32(tail, central + 42));
            long extra = central + 46 + "big.bin".Length;
            Assert.Equal(1, ZipLayout.U16(tail, extra));
            Assert.Equal(16, ZipLayout.U16(tail, extra + 2));
            Assert.Equal((ulong)Size, ZipLayout.U64(tail, extra + 4));
            Assert.Equal((ulong)Size, ZipLayout.U64(tail, extra + 12));

            long second = extra + 20;
            Assert.Equal(0x0201_4B50u, ZipLayout.U32(tail, second));
            Assert.Equal(45, ZipLayout.U16(tail, second + 6));
            Assert.Equal(uint.MaxValue, ZipLayout.U32(tail, second + 42));
            long secondExtra = second + 46 + "after.xml".Length;
            Assert.Equal(1, ZipLayout.U16(tail, secondExtra));
            Assert.Equal(8, ZipLayout.U16(tail, secondExtra + 2));
            Assert.Equal((ulong)(descriptorAt + 24), ZipLayout.U64(tail, secondExtra + 4));
        }

        [Fact]
        public void MoreThan65534EntriesUseZip64EndRecords()
        {
            var ms = new MemoryStream();
            var zip = new ZipPackageWriter(ms);
            for (int i = 0; i < 70_000; i++)
            {
                using Stream entry = zip.OpenEntry($"e{i}.bin", CompressionLevel.NoCompression);
                entry.WriteByte((byte)i);
            }
            zip.Finish();
            byte[] bytes = ms.ToArray();
            int end = bytes.Length - 22;
            Assert.Equal(0xFFFF, ZipLayout.U16(bytes, end + 10));
            int locator = end - 20;
            Assert.Equal(0x0706_4B50u, ZipLayout.U32(bytes, locator));
            long zip64End = (long)ZipLayout.U64(bytes, locator + 8);
            Assert.Equal(0x0606_4B50u, ZipLayout.U32(bytes, zip64End));
            Assert.Equal(70_000UL, ZipLayout.U64(bytes, zip64End + 32));
            using var read = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
            Assert.Equal(70_000, read.Entries.Count);
            using Stream last = read.GetEntry("e69999.bin")!.Open();
            Assert.Equal(69_999 & 0xFF, last.ReadByte());
        }
    }
}
