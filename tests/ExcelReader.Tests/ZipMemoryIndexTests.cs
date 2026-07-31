using System.Buffers.Binary;
using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    // The in-memory ZIP central-directory reader, exercised directly here (rather than only through
    // Excel.From(ReadOnlyMemory<byte>)). Every fixture here is a real
    // ZipArchive-built file, so any divergence from the streamed path (parsed via
    // ZipEntryBytes/ZipArchive in the same test) is a bug in the new reader, not the fixture.
    // No CRC-32 check: ZipArchive doesn't validate it on read either (see ZipMemoryIndex.OpenPart),
    // and it measured at ~48% of a large part's read time for a check the streamed path never paid.
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
            // TryGetEntry returns the first name match, so a second central-directory record
            // sharing a name would silently be unreachable - reject the file outright instead.
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
            // The central directory and local header each carry their own copy of the entry
            // name. Patch only the local header's copy (same length, so no offset shifts) and confirm
            // the mismatch is rejected instead of silently reading the entry under the wrong identity.
            byte[] payload = Encoding.UTF8.GetBytes("hello world, stored and not deflated");
            byte[] zipBytes = BuildZipWithOneEntry("hello.txt", payload, CompressionLevel.NoCompression);

            long localHeaderOffset;
            using (ZipMemoryIndex index = ZipMemoryIndex.Create(zipBytes, ExcelReaderOptions.Default))
            {
                Assert.True(index.TryGetEntry("hello.txt"u8, out ZipEntryRef entry));
                localHeaderOffset = entry.LocalHeaderOffset;
            }

            // Local header layout: 4-byte signature + 26 bytes of fixed fields, then the name.
            int nameOffset = (int)localHeaderOffset + 30;
            Assert.Equal((byte)'h', zipBytes[nameOffset]);
            zipBytes[nameOffset] = (byte)'j'; // "hello.txt" -> "jello.txt" in the local header only

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

        // Mirrors FuzzTests.cs's seeded-mutation harness, aimed squarely at the new untrusted-input
        // boundary (central directory / local header parsing) rather than the row scanners it already
        // covers. Any exception escaping a round that isn't a graceful, documented rejection is a bug.
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
                    // Expected: the mutated bytes were rejected gracefully.
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

        // Hand-crafted ZIP64 EOCD structures, not producible via ZipArchive (which never emits
        // absurd 64-bit offsets), that exercise the two overflow-prone additions fixed in ZipMemoryIndex:
        // ReadZip64Eocd's own `offset` bound, and Create's `cdOffset + cdSize` bound. Every multi-byte
        // field below is little-endian, matching BinaryPrimitives.ReadXLittleEndian on the reader side.

        private const uint Zip64SentinelU32 = 0xFFFFFFFFu;

        private static void WriteEocd(byte[] bytes, int offset, ushort declaredCount, uint cdSize, uint cdOffset)
        {
            BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(offset), 0x06054b50);
            BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(offset + 10), declaredCount);
            BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(offset + 12), cdSize);
            BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(offset + 16), cdOffset);
            // comment length at offset+20 is left 0.
        }

        private static void WriteZip64Locator(byte[] bytes, int offset, long zip64EocdOffset)
        {
            BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(offset), 0x07064b50);
            BinaryPrimitives.WriteInt64LittleEndian(bytes.AsSpan(offset + 8), zip64EocdOffset);
        }

        [Fact]
        public void HugeZip64EocdLocatorOffsetThrowsInsteadOfWrapping()
        {
            // Layout: [0..20) Zip64 locator, pointing to a Zip64 EOCD offset near long.MaxValue,
            // immediately followed by [20..42) a regular EOCD whose sentinel fields force the ZIP64 path.
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
            // Layout: [0..56) a valid, in-range Zip64 EOCD record whose cdSize field is near
            // long.MaxValue (cdOffset stays small), [56..76) the Zip64 locator pointing at it,
            // [76..98) the regular EOCD with ZIP64 sentinel fields.
            byte[] bytes = new byte[98];
            BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(0), 0x06064b50);
            BinaryPrimitives.WriteInt64LittleEndian(bytes.AsSpan(32), 1); // total entries
            BinaryPrimitives.WriteInt64LittleEndian(bytes.AsSpan(40), long.MaxValue - 2); // cdSize
            BinaryPrimitives.WriteInt64LittleEndian(bytes.AsSpan(48), 5); // cdOffset
            WriteZip64Locator(bytes, 56, zip64EocdOffset: 0);
            WriteEocd(bytes, 76, declaredCount: 0xFFFF, cdSize: Zip64SentinelU32, cdOffset: Zip64SentinelU32);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => ZipMemoryIndex.Create(bytes, ExcelReaderOptions.Default));
            Assert.Contains("central directory", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ZipEntryBytesReadThrowsInvalidDataWhenEntryUnderDelivers()
        {
            // The streamed path (ZipEntryBytes.Read, used via a real ZipArchive) previously let
            // a raw EndOfStreamException escape here instead of the Reader-layer's InvalidDataException
            // convention — the in-memory twin (ZipMemoryIndex.InflateToPart) already rewrapped it.
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
            // Attempted an over-delivery check (declared size smaller than what
            // actually decompresses) mirroring ZipMemoryIndex.InflateToPart's trailing ReadByte() check,
            // symmetric with the under-delivery fix below. Verified by direct experiment that
            // ZipArchiveEntry.Open() itself silently stops the stream at the entry's declared Length —
            // the extra real decompressed byte here is simply never exposed, so a "read one more, expect
            // EOF" check always passes and can never catch this. This documents that BCL behavior so it
            // isn't rediscovered as a bug in ZipEntryBytes.Read: over-delivery is not detectable on this
            // path without bypassing ZipArchiveEntry.Open() and driving DeflateStream directly, the way
            // ZipMemoryIndex.InflateToPart already does for the in-memory path.
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
