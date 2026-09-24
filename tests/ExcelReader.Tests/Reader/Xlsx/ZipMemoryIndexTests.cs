using System.Buffers.Binary;
using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Zip;

namespace ExcelReader.Tests.Reader.Xlsx
{
    public class ZipMemoryIndexTests
    {
        [Fact]
        public void OpenPartOnDeflatedEntryMatchesStreamedRead()
        {
            using MemoryStream built = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            byte[] zipBytes = built.ToArray();
            byte[] expected = ReadViaZipArchive(zipBytes, "xl/workbook.xml");

            using ZipMemoryIndex index = ZipMemoryIndex.Create(zipBytes, ExcelReaderOptions.Default);
            Assert.True(index.TryGetEntry("xl/workbook.xml"u8, out ZipEntryRef entry));
            Assert.Equal((ushort)8, entry.Method);

            using ZipPart part = index.OpenPart(entry, new DecompressedByteCounter(0));
            Assert.Equal(expected, part.Memory.ToArray());
        }

        [Fact]
        public void OpenPartOnStoredEntryAliasesTheSourceBufferWithoutCopying()
        {
            byte[] payload = Encoding.UTF8.GetBytes("hello world, stored and not deflated");
            byte[] zipBytes = BuildZipWithOneEntry("hello.txt", payload, CompressionLevel.NoCompression);

            using ZipMemoryIndex index = ZipMemoryIndex.Create(zipBytes, ExcelReaderOptions.Default);
            Assert.True(index.TryGetEntry("hello.txt"u8, out ZipEntryRef entry));
            Assert.Equal((ushort)0, entry.Method);

            using ZipPart part = index.OpenPart(entry, new DecompressedByteCounter(0));
            Assert.Equal(payload, part.Memory.ToArray());
            Assert.True(MemoryMarshal.TryGetArray(part.Memory, out ArraySegment<byte> segment));
            Assert.Same(zipBytes, segment.Array);
        }

        [Fact]
        public void TryGetEntryReturnsFalseForAnAbsentName()
        {
            using MemoryStream built = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            byte[] zipBytes = built.ToArray();

            using ZipMemoryIndex index = ZipMemoryIndex.Create(zipBytes, ExcelReaderOptions.Default);
            Assert.False(index.TryGetEntry("xl/doesNotExist.xml"u8, out _));
        }

        [Fact]
        public void CreateThrowsOnDuplicateCentralDirectoryEntryNames()
        {
            using var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                using (var s1 = zip.CreateEntry("dup.txt", CompressionLevel.NoCompression).Open())
                {
                    s1.Write("first"u8);
                }
                using (var s2 = zip.CreateEntry("dup.txt", CompressionLevel.NoCompression).Open())
                {
                    s2.Write("second"u8);
                }
            }
            byte[] zipBytes = ms.ToArray();

            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => ZipMemoryIndex.Create(zipBytes, ExcelReaderOptions.Default));
            Assert.Contains("duplicate", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void OpenPartThrowsWhenLocalHeaderNameDoesNotMatchCentralDirectory()
        {
            byte[] payload = Encoding.UTF8.GetBytes("hello world, stored and not deflated");
            byte[] zipBytes = BuildZipWithOneEntry("hello.txt", payload, CompressionLevel.NoCompression);

            long localHeaderOffset;
            using (ZipMemoryIndex index = ZipMemoryIndex.Create(zipBytes, ExcelReaderOptions.Default))
            {
                Assert.True(index.TryGetEntry("hello.txt"u8, out ZipEntryRef entry));
                localHeaderOffset = entry.LocalHeaderOffset;
            }

            int nameOffset = (int)localHeaderOffset + 30;
            Assert.Equal((byte)'h', zipBytes[nameOffset]);
            zipBytes[nameOffset] = (byte)'j';

            using ZipMemoryIndex mutatedIndex = ZipMemoryIndex.Create(zipBytes, ExcelReaderOptions.Default);
            Assert.True(mutatedIndex.TryGetEntry("hello.txt"u8, out ZipEntryRef mutatedEntry));
            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => mutatedIndex.OpenPart(mutatedEntry, new DecompressedByteCounter(0)));
            Assert.Contains("does not match", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void OpenPartThrowsWhenDeclaredSizeExceedsTheRemainingBudget()
        {
            using MemoryStream built = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            byte[] zipBytes = built.ToArray();
            PatchCentralDirectoryUInt32(zipBytes, "xl/workbook.xml", fieldOffset: 24, 50_000_000);

            using ZipMemoryIndex index = ZipMemoryIndex.Create(zipBytes, ExcelReaderOptions.Default);
            Assert.True(index.TryGetEntry("xl/workbook.xml"u8, out ZipEntryRef entry));
            var counter = new DecompressedByteCounter(4096, nameof(ExcelReaderOptions.MaxTotalDecompressedBytes));

            ExcelLimitExceededException ex = Assert.Throws<ExcelLimitExceededException>(
                () => index.OpenPart(entry, counter));
            Assert.Equal(nameof(ExcelReaderOptions.MaxTotalDecompressedBytes), ex.LimitName);
        }

        [Fact]
        public void CreateThrowsWhenEntryCountExceedsMaxZipEntries()
        {
            using MemoryStream built = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            byte[] baseBytes = built.ToArray();
            using var ms = new MemoryStream();
            ms.Write(baseBytes, 0, baseBytes.Length);
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
            {
                for (int i = 0; i < 20; i++)
                {
                    using var entryStream = zip.CreateEntry($"junk/{i}.txt").Open();
                }
            }
            byte[] zipBytes = ms.ToArray();
            var options = new ExcelReaderOptions { MaxZipEntries = 10 };

            ExcelLimitExceededException ex = Assert.Throws<ExcelLimitExceededException>(
                () => ZipMemoryIndex.Create(zipBytes, options));
            Assert.Equal(nameof(ExcelReaderOptions.MaxZipEntries), ex.LimitName);
        }

        [Fact]
        public void CreateThrowsForAnEncryptedEntry()
        {
            using MemoryStream built = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            byte[] zipBytes = built.ToArray();
            PatchCentralDirectoryUInt16(zipBytes, "xl/workbook.xml", fieldOffset: 8, 0x0001);

            Assert.Throws<NotSupportedException>(() => ZipMemoryIndex.Create(zipBytes, ExcelReaderOptions.Default));
        }

        [Fact]
        public void OpenPartThrowsForAnUnsupportedCompressionMethod()
        {
            using MemoryStream built = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            byte[] zipBytes = built.ToArray();
            PatchCentralDirectoryUInt16(zipBytes, "xl/workbook.xml", fieldOffset: 10, 12);

            using ZipMemoryIndex index = ZipMemoryIndex.Create(zipBytes, ExcelReaderOptions.Default);
            Assert.True(index.TryGetEntry("xl/workbook.xml"u8, out ZipEntryRef entry));
            Assert.Throws<NotSupportedException>(() => index.OpenPart(entry, new DecompressedByteCounter(0)));
        }

        [Fact]
        public void CreateThrowsWhenNoEndOfCentralDirectoryRecordExists()
        {
            byte[] notAZip = Encoding.UTF8.GetBytes("this is not a zip file at all");
            Assert.Throws<InvalidDataException>(() => ZipMemoryIndex.Create(notAZip, ExcelReaderOptions.Default));
        }

        [Fact]
        public void MutatedZipBytesNeverCrashTheMemoryIndex()
        {
            using MemoryStream built = WorkbookBuilder.BuildMultiSheet(
                sheets:
                [
                    ("S1", """<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row>"""),
                    ("S2", """<row r="1"><c r="A1"><v>7</v></c></row>"""),
                ],
                sharedStrings: "<si><t>hello</t></si><si><t>world</t></si>",
                styles: "<styleSheet><cellXfs count=\"1\"><xf numFmtId=\"14\"/></cellXfs></styleSheet>");
            byte[] seed = built.ToArray();

            var rng = new Random(20260728);
            const int rounds = 500;
            int completed = 0;
            for (int round = 0; round < rounds; round++)
            {
                byte[] mutated = FuzzMutation.MutateCopy(seed, rng, out int[] positions);
                try
                {
                    FuzzMutation.RunBounded(() => OpenAllPartsAndDrain(mutated));
                }
                catch (Exception ex) when (FuzzMutation.IsAcceptable(ex))
                {
                }
                catch (Exception ex)
                {
                    string offsets = string.Join(", ", positions);
                    throw new InvalidOperationException(
                        string.Create(CultureInfo.InvariantCulture,
                            $"Round {round} produced an unhandled '{ex.GetType().Name}' (mutated byte offsets: [{offsets}])."),
                        ex);
                }
                completed++;
            }
            Assert.Equal(rounds, completed);
        }

        private static void OpenAllPartsAndDrain(byte[] bytes)
        {
            using ZipMemoryIndex index = ZipMemoryIndex.Create(bytes, ExcelReaderOptions.Default);
            var counter = new DecompressedByteCounter(ExcelReaderOptions.Default.MaxTotalDecompressedBytes);
            foreach (string name in (string[])["xl/workbook.xml", "xl/sharedStrings.xml", "xl/styles.xml", "xl/worksheets/sheet1.xml"])
            {
                if (!index.TryGetEntry(Encoding.UTF8.GetBytes(name), out ZipEntryRef entry))
                {
                    continue;
                }
                using ZipPart part = index.OpenPart(entry, counter);
                _ = part.Memory.Length;
            }
        }


        private const uint Zip64SentinelU32 = 0xFFFFFFFFu;

        private static void WriteEocd(byte[] bytes, int offset, ushort declaredCount, uint cdSize, uint cdOffset)
        {
            BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(offset), 0x06054b50);
            BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(offset + 10), declaredCount);
            BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(offset + 12), cdSize);
            BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(offset + 16), cdOffset);
        }

        private static void WriteZip64Locator(byte[] bytes, int offset, long zip64EocdOffset)
        {
            BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(offset), 0x07064b50);
            BinaryPrimitives.WriteInt64LittleEndian(bytes.AsSpan(offset + 8), zip64EocdOffset);
        }

        [Fact]
        public void HugeZip64EocdLocatorOffsetThrowsInsteadOfWrapping()
        {
            byte[] bytes = new byte[42];
            WriteZip64Locator(bytes, 0, zip64EocdOffset: long.MaxValue - 2);
            WriteEocd(bytes, 20, declaredCount: 0xFFFF, cdSize: Zip64SentinelU32, cdOffset: Zip64SentinelU32);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => ZipMemoryIndex.Create(bytes, ExcelReaderOptions.Default));
            Assert.Contains("ZIP64", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void HugeZip64CentralDirectorySizeThrowsInsteadOfWrapping()
        {
            byte[] bytes = new byte[98];
            BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(0), 0x06064b50);
            BinaryPrimitives.WriteInt64LittleEndian(bytes.AsSpan(32), 1);
            BinaryPrimitives.WriteInt64LittleEndian(bytes.AsSpan(40), long.MaxValue - 2);
            BinaryPrimitives.WriteInt64LittleEndian(bytes.AsSpan(48), 5);
            WriteZip64Locator(bytes, 56, zip64EocdOffset: 0);
            WriteEocd(bytes, 76, declaredCount: 0xFFFF, cdSize: Zip64SentinelU32, cdOffset: Zip64SentinelU32);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => ZipMemoryIndex.Create(bytes, ExcelReaderOptions.Default));
            Assert.Contains("central directory", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        private static byte[] BuildCdWithZip64Field(uint compressedSize, uint uncompressedSize, long zip64Value)
        {
            byte[] name = "a"u8.ToArray();
            byte[] extraData = new byte[8];
            BinaryPrimitives.WriteInt64LittleEndian(extraData, zip64Value);
            ushort extraLength = (ushort)(4 + extraData.Length);
            int cdSize = CentralDirectoryFixedSize + name.Length + extraLength;
            byte[] bytes = new byte[cdSize + 22];

            BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(0), 0x02014b50);
            BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(20), compressedSize);
            BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(24), uncompressedSize);
            BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(28), (ushort)name.Length);
            BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(30), extraLength);

            int nameOffset = CentralDirectoryFixedSize;
            name.CopyTo(bytes.AsSpan(nameOffset));
            int extraOffset = nameOffset + name.Length;
            BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(extraOffset), 0x0001);
            BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(extraOffset + 2), (ushort)extraData.Length);
            extraData.CopyTo(bytes.AsSpan(extraOffset + 4));

            WriteEocd(bytes, cdSize, declaredCount: 1, cdSize: (uint)cdSize, cdOffset: 0);
            return bytes;
        }

        private const int CentralDirectoryFixedSize = 46;

        [Fact]
        public void NegativeZip64UncompressedSizeThrowsInvalidDataException()
        {
            byte[] bytes = BuildCdWithZip64Field(compressedSize: 5, uncompressedSize: Zip64SentinelU32, zip64Value: -1);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => ZipMemoryIndex.Create(bytes, ExcelReaderOptions.Default));
            Assert.Contains("negative", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void NegativeZip64CompressedSizeThrowsInvalidDataException()
        {
            byte[] bytes = BuildCdWithZip64Field(compressedSize: Zip64SentinelU32, uncompressedSize: 5, zip64Value: -1);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => ZipMemoryIndex.Create(bytes, ExcelReaderOptions.Default));
            Assert.Contains("negative", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        private static readonly byte[] Zip64Payload = Encoding.UTF8.GetBytes("zip64 stored payload");

        [Fact]
        public void ValidZip64ArchiveOpensTheStoredEntry()
        {
            Zip64Layout zip = BuildStoredZip64(Zip64Payload);

            using ZipMemoryIndex index = ZipMemoryIndex.Create(zip.Bytes, ExcelReaderOptions.Default);
            Assert.True(index.TryGetEntry("a.txt"u8, out ZipEntryRef entry));
            Assert.Equal(0, entry.LocalHeaderOffset);
            Assert.Equal(Zip64Payload.Length, entry.CompressedSize);
            Assert.Equal(Zip64Payload.Length, entry.UncompressedSize);
            using ZipPart part = index.OpenPart(entry, new DecompressedByteCounter(0));
            Assert.Equal(Zip64Payload, part.Memory.ToArray());

            using var ms = new MemoryStream(zip.Bytes);
            using var bcl = new ZipArchive(ms, ZipArchiveMode.Read);
            using Stream bclStream = bcl.GetEntry("a.txt")!.Open();
            using var copy = new MemoryStream();
            bclStream.CopyTo(copy);
            Assert.Equal(Zip64Payload, copy.ToArray());
        }

        [Theory]
        [InlineData(long.MaxValue)]
        [InlineData(long.MaxValue - 29)]
        [InlineData(long.MaxValue - 30)]
        [InlineData((long)int.MaxValue + 1)]
        [InlineData(int.MaxValue)]
        public void HugeZip64LocalHeaderOffsetThrowsInvalidDataOnOpen(long localOffset)
        {
            Zip64Layout zip = BuildStoredZip64(Zip64Payload, localOffset: localOffset);

            using ZipMemoryIndex index = ZipMemoryIndex.Create(zip.Bytes, ExcelReaderOptions.Default);
            Assert.True(index.TryGetEntry("a.txt"u8, out ZipEntryRef entry));
            Assert.Equal(localOffset, entry.LocalHeaderOffset);
            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => index.OpenPart(entry, new DecompressedByteCounter(0)));
            Assert.Contains("local file header", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void LocalHeaderOffsetJustInsideTheFileEndThrowsInvalidDataOnOpen()
        {
            Zip64Layout probe = BuildStoredZip64(Zip64Payload);
            Zip64Layout zip = BuildStoredZip64(Zip64Payload, localOffset: probe.Bytes.Length - 29);

            using ZipMemoryIndex index = ZipMemoryIndex.Create(zip.Bytes, ExcelReaderOptions.Default);
            Assert.True(index.TryGetEntry("a.txt"u8, out ZipEntryRef entry));
            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => index.OpenPart(entry, new DecompressedByteCounter(0)));
            Assert.Contains("out of range", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Theory]
        [InlineData(long.MaxValue)]
        [InlineData(long.MaxValue - 50)]
        [InlineData((long)int.MaxValue + 1)]
        [InlineData(int.MaxValue)]
        public void HugeZip64CompressedSizeThrowsInvalidDataOnOpen(long compressedSize)
        {
            Zip64Layout zip = BuildStoredZip64(Zip64Payload, compressedSize: compressedSize);

            using ZipMemoryIndex index = ZipMemoryIndex.Create(zip.Bytes, ExcelReaderOptions.Default);
            Assert.True(index.TryGetEntry("a.txt"u8, out ZipEntryRef entry));
            Assert.Equal(compressedSize, entry.CompressedSize);
            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => index.OpenPart(entry, new DecompressedByteCounter(0)));
            Assert.Contains("past the end", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Throws<InvalidDataException>(
                () => index.OpenEntryStream(entry, new DecompressedByteCounter(0), ExcelReaderOptions.Default));
        }

        [Fact]
        public void CompressedSizeEndingExactlyAtTheFileEndIsAcceptedAndOneMoreIsRejected()
        {
            Zip64Layout probe = BuildStoredZip64(Zip64Payload);
            long exact = probe.Bytes.Length - probe.DataOffset;

            Zip64Layout fits = BuildStoredZip64(Zip64Payload, compressedSize: exact);
            using (ZipMemoryIndex index = ZipMemoryIndex.Create(fits.Bytes, ExcelReaderOptions.Default))
            {
                Assert.True(index.TryGetEntry("a.txt"u8, out ZipEntryRef entry));
                using ZipPart part = index.OpenPart(entry, new DecompressedByteCounter(0));
                Assert.Equal(exact, part.Memory.Length);
            }

            Zip64Layout overruns = BuildStoredZip64(Zip64Payload, compressedSize: exact + 1);
            using (ZipMemoryIndex index = ZipMemoryIndex.Create(overruns.Bytes, ExcelReaderOptions.Default))
            {
                Assert.True(index.TryGetEntry("a.txt"u8, out ZipEntryRef entry));
                Assert.Throws<InvalidDataException>(() => index.OpenPart(entry, new DecompressedByteCounter(0)));
            }
        }

        [Theory]
        [InlineData(0x8000000000000000UL)]
        [InlineData(ulong.MaxValue)]
        public void Zip64LocalHeaderOffsetAboveLongMaxIsRejectedAsNegative(ulong localOffset)
        {
            Zip64Layout zip = BuildStoredZip64(Zip64Payload, localOffset: unchecked((long)localOffset));

            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => ZipMemoryIndex.Create(zip.Bytes, ExcelReaderOptions.Default));
            Assert.Contains("negative", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Theory]
        [InlineData(0x8000000000000000UL)]
        [InlineData(ulong.MaxValue)]
        public void Zip64LocatorOffsetAboveLongMaxIsRejectedNotIgnored(ulong zip64EocdOffset)
        {
            Zip64Layout zip = BuildStoredZip64(Zip64Payload);
            BinaryPrimitives.WriteUInt64LittleEndian(zip.Bytes.AsSpan(zip.LocatorOffset + 8), zip64EocdOffset);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => ZipMemoryIndex.Create(zip.Bytes, ExcelReaderOptions.Default));
            Assert.Contains("ZIP64", ex.Message, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(long.MaxValue, 1UL)]
        [InlineData(long.MaxValue, (ulong)long.MaxValue)]
        [InlineData(1L, (ulong)long.MaxValue)]
        [InlineData(0L, ulong.MaxValue)]
        [InlineData(-1L, 0UL)]
        [InlineData(long.MinValue, 100UL)]
        [InlineData(0L, 0x8000000000000000UL)]
        public void Zip64CentralDirectoryBoundsThatWouldOverflowAreRejected(long cdOffset, ulong cdSize)
        {
            Zip64Layout zip = BuildStoredZip64(Zip64Payload);
            BinaryPrimitives.WriteUInt64LittleEndian(zip.Bytes.AsSpan(zip.Zip64EocdOffset + 40), cdSize);
            BinaryPrimitives.WriteInt64LittleEndian(zip.Bytes.AsSpan(zip.Zip64EocdOffset + 48), cdOffset);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => ZipMemoryIndex.Create(zip.Bytes, ExcelReaderOptions.Default));
            Assert.Contains("central directory", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Zip64CentralDirectorySizeOneByteLongerThanItsRecordsIsRejected()
        {
            Zip64Layout zip = BuildStoredZip64(Zip64Payload);
            long cdOffset = BinaryPrimitives.ReadInt64LittleEndian(zip.Bytes.AsSpan(zip.Zip64EocdOffset + 48));
            long cdSize = BinaryPrimitives.ReadInt64LittleEndian(zip.Bytes.AsSpan(zip.Zip64EocdOffset + 40));
            Assert.Equal(zip.Zip64EocdOffset, cdOffset + cdSize);

            BinaryPrimitives.WriteInt64LittleEndian(zip.Bytes.AsSpan(zip.Zip64EocdOffset + 40), cdSize + 1);
            Assert.Throws<InvalidDataException>(() => ZipMemoryIndex.Create(zip.Bytes, ExcelReaderOptions.Default));
        }

        [Theory]
        [InlineData(28, (ushort)10)]
        [InlineData(30, (ushort)0xFFFF)]
        [InlineData(32, (ushort)1)]
        [InlineData(32, (ushort)0xFFFF)]
        public void CentralDirectoryVariableFieldsRunningPastTheDirectoryEndAreRejected(int fieldOffset, ushort value)
        {
            byte[] zipBytes = BuildZipWithOneEntry("hello.txt", Zip64Payload, CompressionLevel.NoCompression);
            PatchCentralDirectoryUInt16(zipBytes, "hello.txt", fieldOffset, value);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => ZipMemoryIndex.Create(zipBytes, ExcelReaderOptions.Default));
            Assert.Contains("truncated", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Theory]
        [InlineData(1u)]
        [InlineData(45u)]
        public void CentralDirectorySizeShorterThanItsFixedRecordIsRejected(uint cdSize)
        {
            byte[] zipBytes = BuildZipWithOneEntry("hello.txt", Zip64Payload, CompressionLevel.NoCompression);
            int eocd = zipBytes.Length - 22;
            uint cdOffset = BinaryPrimitives.ReadUInt32LittleEndian(zipBytes.AsSpan(eocd + 16));
            WriteEocd(zipBytes, eocd, declaredCount: 1, cdSize: cdSize, cdOffset: cdOffset);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => ZipMemoryIndex.Create(zipBytes, ExcelReaderOptions.Default));
            Assert.Contains("truncated", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void LocalExtraLengthPushingDataPastTheFileEndThrowsInvalidDataOnOpen()
        {
            byte[] zipBytes = BuildZipWithOneEntry("hello.txt", Zip64Payload, CompressionLevel.NoCompression);
            BinaryPrimitives.WriteUInt16LittleEndian(zipBytes.AsSpan(28), 0xFFFF);

            using ZipMemoryIndex index = ZipMemoryIndex.Create(zipBytes, ExcelReaderOptions.Default);
            Assert.True(index.TryGetEntry("hello.txt"u8, out ZipEntryRef entry));
            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => index.OpenPart(entry, new DecompressedByteCounter(0)));
            Assert.Contains("past the end", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        private readonly record struct Zip64Layout(byte[] Bytes, int DataOffset, int Zip64EocdOffset, int LocatorOffset);

        private static Zip64Layout BuildStoredZip64(byte[] payload, long? localOffset = null, long? compressedSize = null)
        {
            byte[] name = "a.txt"u8.ToArray();
            uint crc = Crc32(payload);
            int localExtraLength = 4 + 16;
            int cdExtraLength = 4 + 24;
            int dataOffset = 30 + name.Length + localExtraLength;
            int cdOffset = dataOffset + payload.Length;
            int cdSize = CentralDirectoryFixedSize + name.Length + cdExtraLength;
            int zip64EocdOffset = cdOffset + cdSize;
            int locatorOffset = zip64EocdOffset + 56;
            byte[] bytes = new byte[locatorOffset + 20 + 22];
            Span<byte> s = bytes;

            BinaryPrimitives.WriteInt32LittleEndian(s, 0x04034b50);
            BinaryPrimitives.WriteUInt16LittleEndian(s[4..], 45);
            BinaryPrimitives.WriteUInt32LittleEndian(s[14..], crc);
            BinaryPrimitives.WriteUInt32LittleEndian(s[18..], Zip64SentinelU32);
            BinaryPrimitives.WriteUInt32LittleEndian(s[22..], Zip64SentinelU32);
            BinaryPrimitives.WriteUInt16LittleEndian(s[26..], (ushort)name.Length);
            BinaryPrimitives.WriteUInt16LittleEndian(s[28..], (ushort)localExtraLength);
            name.CopyTo(s[30..]);
            WriteZip64Extra(s[(30 + name.Length)..], payload.Length, payload.Length);
            payload.CopyTo(s[dataOffset..]);

            Span<byte> cd = s[cdOffset..];
            BinaryPrimitives.WriteInt32LittleEndian(cd, 0x02014b50);
            BinaryPrimitives.WriteUInt16LittleEndian(cd[4..], 45);
            BinaryPrimitives.WriteUInt16LittleEndian(cd[6..], 45);
            BinaryPrimitives.WriteUInt32LittleEndian(cd[16..], crc);
            BinaryPrimitives.WriteUInt32LittleEndian(cd[20..], Zip64SentinelU32);
            BinaryPrimitives.WriteUInt32LittleEndian(cd[24..], Zip64SentinelU32);
            BinaryPrimitives.WriteUInt16LittleEndian(cd[28..], (ushort)name.Length);
            BinaryPrimitives.WriteUInt16LittleEndian(cd[30..], (ushort)cdExtraLength);
            BinaryPrimitives.WriteUInt32LittleEndian(cd[42..], Zip64SentinelU32);
            name.CopyTo(cd[CentralDirectoryFixedSize..]);
            WriteZip64Extra(cd[(CentralDirectoryFixedSize + name.Length)..],
                payload.Length, compressedSize ?? payload.Length, localOffset ?? 0);

            Span<byte> z64 = s[zip64EocdOffset..];
            BinaryPrimitives.WriteInt32LittleEndian(z64, 0x06064b50);
            BinaryPrimitives.WriteInt64LittleEndian(z64[4..], 44);
            BinaryPrimitives.WriteUInt16LittleEndian(z64[12..], 45);
            BinaryPrimitives.WriteUInt16LittleEndian(z64[14..], 45);
            BinaryPrimitives.WriteInt64LittleEndian(z64[24..], 1);
            BinaryPrimitives.WriteInt64LittleEndian(z64[32..], 1);
            BinaryPrimitives.WriteInt64LittleEndian(z64[40..], cdSize);
            BinaryPrimitives.WriteInt64LittleEndian(z64[48..], cdOffset);

            WriteZip64Locator(bytes, locatorOffset, zip64EocdOffset);
            BinaryPrimitives.WriteUInt32LittleEndian(s[(locatorOffset + 16)..], 1);
            int eocd = locatorOffset + 20;
            WriteEocd(bytes, eocd, declaredCount: 0xFFFF, cdSize: Zip64SentinelU32, cdOffset: Zip64SentinelU32);
            BinaryPrimitives.WriteUInt16LittleEndian(s[(eocd + 8)..], 0xFFFF);
            return new Zip64Layout(bytes, dataOffset, zip64EocdOffset, locatorOffset);
        }

        private static void WriteZip64Extra(Span<byte> destination, params ReadOnlySpan<long> values)
        {
            BinaryPrimitives.WriteUInt16LittleEndian(destination, 0x0001);
            BinaryPrimitives.WriteUInt16LittleEndian(destination[2..], (ushort)(values.Length * 8));
            for (int i = 0; i < values.Length; i++)
            {
                BinaryPrimitives.WriteInt64LittleEndian(destination[(4 + (i * 8))..], values[i]);
            }
        }

        private static uint Crc32(ReadOnlySpan<byte> data)
        {
            uint crc = 0xFFFFFFFFu;
            foreach (byte b in data)
            {
                crc ^= b;
                for (int k = 0; k < 8; k++)
                {
                    crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320u : crc >> 1;
                }
            }
            return ~crc;
        }

        [Fact]
        public void ZipEntryBytesReadThrowsInvalidDataWhenEntryUnderDelivers()
        {
            using MemoryStream built = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            byte[] zipBytes = built.ToArray();
            uint realLength = ReadDeclaredUncompressedSize(zipBytes, "xl/workbook.xml");
            PatchCentralDirectoryUInt32(zipBytes, "xl/workbook.xml", fieldOffset: 24, realLength + 64);

            using var ms = new MemoryStream(zipBytes);
            using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => ZipEntryBytes.Read(zip, "xl/workbook.xml", new DecompressedByteCounter(0)));
            Assert.Contains("less data", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ZipArchiveEntryOpenSilentlyTruncatesAtDeclaredLengthUnderOverDelivery()
        {
            using MemoryStream built = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            byte[] zipBytes = built.ToArray();
            uint realLength = ReadDeclaredUncompressedSize(zipBytes, "xl/workbook.xml");
            Assert.True(realLength > 0);
            PatchCentralDirectoryUInt32(zipBytes, "xl/workbook.xml", fieldOffset: 24, realLength - 1);

            using var ms = new MemoryStream(zipBytes);
            using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
            ZipArchiveEntry entry = zip.GetEntry("xl/workbook.xml")!;
            using Stream stream = entry.Open();
            byte[] buffer = new byte[(int)entry.Length];
            stream.ReadExactly(buffer);
            Assert.Equal(-1, stream.ReadByte());
        }

        private static uint ReadDeclaredUncompressedSize(byte[] zipBytes, string entryName)
        {
            int cdOffset = FindCentralDirectoryOffset(zipBytes, entryName);
            return BinaryPrimitives.ReadUInt32LittleEndian(zipBytes.AsSpan(cdOffset + 24, 4));
        }

        private static byte[] ReadViaZipArchive(byte[] zipBytes, string entryName)
        {
            using var ms = new MemoryStream(zipBytes);
            using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
            using ZipPart part = ZipEntryBytes.Read(zip, entryName, new DecompressedByteCounter(0));
            return part.Memory.ToArray();
        }

        private static byte[] BuildZipWithOneEntry(string entryName, byte[] payload, CompressionLevel level)
        {
            using var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                using var entryStream = zip.CreateEntry(entryName, level).Open();
                entryStream.Write(payload, 0, payload.Length);
            }
            return ms.ToArray();
        }

        private static int FindCentralDirectoryOffset(byte[] zipBytes, string entryName)
        {
            byte[] nameBytes = Encoding.UTF8.GetBytes(entryName);
            for (int i = 0; i + 46 <= zipBytes.Length; i++)
            {
                if (!IsCentralDirectoryRecordFor(zipBytes, i, nameBytes))
                {
                    continue;
                }
                return i;
            }
            throw new InvalidOperationException($"Central directory entry '{entryName}' not found.");
        }

        private static bool IsCentralDirectoryRecordFor(byte[] zipBytes, int offset, byte[] nameBytes)
        {
            if (zipBytes[offset] != 0x50 || zipBytes[offset + 1] != 0x4B || zipBytes[offset + 2] != 0x01 || zipBytes[offset + 3] != 0x02)
            {
                return false;
            }
            int nameLen = BinaryPrimitives.ReadUInt16LittleEndian(zipBytes.AsSpan(offset + 28, 2));
            if (nameLen != nameBytes.Length || offset + 46 + nameLen > zipBytes.Length)
            {
                return false;
            }
            return zipBytes.AsSpan(offset + 46, nameLen).SequenceEqual(nameBytes);
        }

        private static void PatchCentralDirectoryUInt32(byte[] zipBytes, string entryName, int fieldOffset, uint value)
        {
            int cdOffset = FindCentralDirectoryOffset(zipBytes, entryName);
            BinaryPrimitives.WriteUInt32LittleEndian(zipBytes.AsSpan(cdOffset + fieldOffset, 4), value);
        }

        private static void PatchCentralDirectoryUInt16(byte[] zipBytes, string entryName, int fieldOffset, ushort value)
        {
            int cdOffset = FindCentralDirectoryOffset(zipBytes, entryName);
            BinaryPrimitives.WriteUInt16LittleEndian(zipBytes.AsSpan(cdOffset + fieldOffset, 2), value);
        }
    }
}
