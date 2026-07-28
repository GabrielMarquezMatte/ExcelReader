using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    // Z1 in docs/in-memory-zip.md: the in-memory ZIP central-directory reader, exercised directly
    // (Excel.From(ReadOnlyMemory<byte>) — Z4 — doesn't exist yet). Every fixture here is a real
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
                byte[] mutated = MutateCopy(seed, rng, out int[] positions);
                try
                {
                    OpenAllPartsAndDrain(mutated);
                }
                catch (Exception ex) when (IsAcceptable(ex))
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

        private static readonly Type[] AcceptableExceptionTypes =
        [
            typeof(InvalidDataException),
            typeof(ExcelLimitExceededException),
            typeof(EndOfStreamException),
            typeof(IOException),
            typeof(OverflowException),
            typeof(ArgumentException),
            typeof(NotSupportedException),
        ];

        private static bool IsAcceptable(Exception ex)
        {
            foreach (Type acceptableType in AcceptableExceptionTypes)
            {
                if (acceptableType.IsInstanceOfType(ex))
                {
                    return true;
                }
            }
            return false;
        }

        [SuppressMessage("Security", "CA5394:Do not use insecure randomness",
            Justification = "Fuzzing needs a reproducible seeded PRNG, not cryptographic randomness.")]
        [SuppressMessage("Performance", "HLQ013:Consider using 'foreach' loop instead of 'for' loop",
            Justification = "Each iteration both reads (rng.Next) and writes positions[i] by index; foreach can't express the write.")]
        private static byte[] MutateCopy(byte[] seed, Random rng, out int[] positions)
        {
            byte[] copy = (byte[])seed.Clone();
            int count = rng.Next(1, 9);
            positions = new int[count];
            for (int i = 0; i < count; i++)
            {
                int pos = rng.Next(copy.Length);
                positions[i] = pos;
                copy[pos] = (byte)rng.Next(256);
            }
            return copy;
        }

        private static byte[] ReadViaZipArchive(byte[] zipBytes, string entryName)
        {
            using var ms = new MemoryStream(zipBytes);
            using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
            return ZipEntryBytes.Read(zip, entryName, new DecompressedByteCounter(0));
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
