using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    public class ReaderLimitTests
    {
        // Regression: a corrupted/malicious file can encode an arbitrary column index in a per-cell
        // record (e.g. a raw 4-byte BIFF12 column field), which used to flow straight into
        // Row.ColumnCount unchecked — turning a fuzzed single-byte flip into a near-infinite loop for
        // any caller iterating 0..ColumnCount, instead of a graceful rejection. CellAccumulator.Add is
        // the one choke point shared by every reader (XLS/XLSB/XLSX/CSV), so the bound lives there.
        [Fact]
        public void CellAccumulatorRejectsColumnIndexAtOrAboveExcelLimit()
        {
            var acc = new CellAccumulator(maxCellBytes: 0, limitName: "Test");
            try
            {
                ExcelLimitExceededException ex = Assert.Throws<ExcelLimitExceededException>(
                    () => acc.Add(16_384, start: 0, len: 0, CellType.ExcelString, style: 0, CellValueSource.RowValues));
                Assert.Equal("Columns", ex.LimitName);
            }
            finally
            {
                acc.Return();
            }
        }

        [Fact]
        public void CellAccumulatorAcceptsColumnIndexAtExcelLimitBoundary()
        {
            var acc = new CellAccumulator(maxCellBytes: 0, limitName: "Test");
            try
            {
                acc.Add(16_383, start: 0, len: 0, CellType.ExcelString, style: 0, CellValueSource.RowValues);
                Assert.Equal(1, acc.Count);
            }
            finally
            {
                acc.Return();
            }
        }
        // Patches the uncompressed-size field of a central-directory record in place, so the entry's
        // declared size (what ZipArchiveEntry.Length reports) lies far above its real, tiny compressed
        // content — the exact shape of a zip-bomb-style amplification attack.
        private static void ForgeCentralDirectoryUncompressedSize(byte[] zipBytes, string entryName, uint forgedSize)
        {
            byte[] nameBytes = Encoding.UTF8.GetBytes(entryName);
            for (int i = 0; i + 46 <= zipBytes.Length; i++)
            {
                if (zipBytes[i] != 0x50 || zipBytes[i + 1] != 0x4B || zipBytes[i + 2] != 0x01 || zipBytes[i + 3] != 0x02)
                {
                    continue;
                }
                int nameLen = BitConverter.ToUInt16(zipBytes, i + 28);
                if (nameLen != nameBytes.Length || i + 46 + nameLen > zipBytes.Length)
                {
                    continue;
                }
                if (!zipBytes.AsSpan(i + 46, nameLen).SequenceEqual(nameBytes))
                {
                    continue;
                }
                BitConverter.GetBytes(forgedSize).CopyTo(zipBytes, i + 24);
                return;
            }
            throw new InvalidOperationException($"Central directory entry '{entryName}' not found.");
        }

        [Fact]
        public void ForgedOversizedStylesEntryTripsTotalLimitBeforeReading()
        {
            using MemoryStream built = WorkbookBuilder.Build(
                """<row r="1"><c r="A1"><v>1</v></c></row>""",
                styles: """<styleSheet><cellXfs count="1"><xf/></cellXfs></styleSheet>""");
            byte[] zipBytes = built.ToArray();
            ForgeCentralDirectoryUncompressedSize(zipBytes, "xl/styles.xml", 50_000_000);
            using var forged = new MemoryStream(zipBytes);

            var options = new ExcelReaderOptions { MaxTotalDecompressedBytes = 4096 };

            ExcelLimitExceededException ex = Assert.Throws<ExcelLimitExceededException>(() =>
            {
                using XlsxReader reader = Excel.From(forged, options: options);
            });
            Assert.Equal(nameof(ExcelReaderOptions.MaxTotalDecompressedBytes), ex.LimitName);
            Assert.Equal(50_000_000, ex.Actual);
        }

        [Fact]
        public void ForgedOversizedSharedStringsEntryTripsSharedStringLimitBeforeReading()
        {
            using MemoryStream built = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v>0</v></c></row>""",
                sharedStrings: "<si><t>x</t></si>");
            byte[] zipBytes = built.ToArray();
            ForgeCentralDirectoryUncompressedSize(zipBytes, "xl/sharedStrings.xml", 50_000_000);
            using var forged = new MemoryStream(zipBytes);

            var options = new ExcelReaderOptions
            {
                MaxTotalDecompressedBytes = 512L * 1024 * 1024,
                MaxSharedStringBytes = 1024,
            };

            using XlsxReader reader = Excel.From(forged, options: options);
            ExcelLimitExceededException ex = Assert.Throws<ExcelLimitExceededException>(() =>
            {
                using XlsxReader.Enumerator e = reader.GetEnumerator();
            });
            Assert.Equal(nameof(ExcelReaderOptions.MaxSharedStringBytes), ex.LimitName);
            Assert.Equal(50_000_000, ex.Actual);
        }

        // SEC-1: <sst uniqueCount="…"> is attacker-controlled independent of the part's real byte
        // length — unlike ForgedOversizedSharedStringsEntryTripsSharedStringLimitBeforeReading above,
        // this entry's *declared central-directory size* is honest and tiny; only the XML attribute
        // lies. Before the fix, `new int[uniqueCount + 1]` sized the offsets array straight from this
        // attribute, so a ~100-byte part could force an allocation many times its own MaxSharedStringBytes
        // budget. LimitChecks.ThrowIfSharedStringCountImplausible now rejects a count the part could not
        // physically contain before that allocation happens.
        [Fact]
        public void ImplausibleUniqueCountTripsSharedStringLimitBeforeAllocating()
        {
            using MemoryStream built = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v>0</v></c></row>""",
                sharedStrings: "<si><t>a</t></si>");
            byte[] zipBytes = built.ToArray();
            using var ms = new MemoryStream();
            ms.Write(zipBytes, 0, zipBytes.Length);
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
            {
                zip.GetEntry("xl/sharedStrings.xml")!.Delete();
                using StreamWriter writer = new(zip.CreateEntry("xl/sharedStrings.xml").Open(), Encoding.UTF8);
                writer.Write(
                    """<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="500000000"><si><t>a</t></si></sst>""");
            }
            ms.Position = 0;

            using XlsxReader reader = Excel.From(ms);
            ExcelLimitExceededException ex = Assert.Throws<ExcelLimitExceededException>(() =>
            {
                using XlsxReader.Enumerator e = reader.GetEnumerator();
            });
            Assert.Equal(nameof(ExcelReaderOptions.MaxSharedStringBytes), ex.LimitName);
            Assert.Equal(500_000_000, ex.Actual);
        }

        // SEC-2: BoundSheet8.lbPlyPos is read as a raw signed int32 with no validation before it becomes
        // a BiffCursor.Position assignment. On the Chained WorkbookStream kind, a negative position used
        // to resolve to a valid-looking byte range elsewhere in the file (the OLE header/preceding
        // sectors) and silently decode wrong bytes as BIFF records — no exception, wrong data — while
        // the Streamed/Contiguous kinds already threw. The fix validates lbPlyPos once, in
        // ParseWorkbookGlobals, before OpenCursor is ever called with it — which runs identically
        // regardless of which WorkbookStream kind BuildWorkbook chose, so this is structurally the same
        // fix for all three kinds rather than three separate ones. Exercised here through the
        // Stream-based open (Streamed/Contiguous, depending on the forged workbook's size) and the
        // ReadOnlyMemory-based open (Contiguous for a small workbook); a dedicated large-workbook fixture
        // to force the Chained kind specifically would strengthen this further but wasn't built here.
        private static void PatchBoundSheetLbPlyPos(byte[] bytes, int forgedOffset)
        {
            for (int i = 0; i + 8 <= bytes.Length; i++)
            {
                if (bytes[i] == 0x85 && bytes[i + 1] == 0x00)
                {
                    BitConverter.GetBytes(forgedOffset).CopyTo(bytes, i + 4);
                    return;
                }
            }
            throw new InvalidOperationException("BoundSheet8 record (0x0085) not found.");
        }

        [Fact]
        public void NegativeBoundSheetOffsetThrowsInsteadOfReadingOutOfRange()
        {
            using MemoryStream built = XlsWorkbookBuilder.Build(sheets: [("S1", [["Alice", 1, true]])]);
            byte[] streamBytes = built.ToArray();
            PatchBoundSheetLbPlyPos(streamBytes, -100);

            InvalidDataException streamEx = Assert.Throws<InvalidDataException>(
                () => Excel.FromXls(new MemoryStream(streamBytes)));

            byte[] memoryBytes = (byte[])streamBytes.Clone();
            InvalidDataException memoryEx = Assert.Throws<InvalidDataException>(
                () => Excel.FromXls(memoryBytes.AsMemory()));

            Assert.Equal(streamEx.Message, memoryEx.Message);
        }

        // SEC-3: the FAT sector immediately follows the 512-byte OLE header (XlsWorkbookBuilder.SectorSize),
        // so sector index N sits at absolute byte offset (N + 1) * 512 — mirrors XlsCompoundFile.SectorOffset.
        // Sector 0 is the FAT sector itself (marked FatSector by the builder); sector 1 is the directory.
        // Overwriting both entries with each other's index turns the FAT into a 2-sector cycle.
        private static void PatchFatCycle(byte[] oleBytes, int sectorA, int sectorB)
        {
            const int SectorSize = 512;
            const int FatSectorOffset = SectorSize; // FAT sector sits right after the header
            BinaryPrimitives.WriteInt32LittleEndian(oleBytes.AsSpan(FatSectorOffset + (sectorA * 4)), sectorB);
            BinaryPrimitives.WriteInt32LittleEndian(oleBytes.AsSpan(FatSectorOffset + (sectorB * 4)), sectorA);
        }

        [Fact]
        public void FatChainCycleThrowsInsteadOfUnboundedAllocation()
        {
            using MemoryStream built = XlsWorkbookBuilder.Build(sheets: [("S1", [["Alice", 1, true]])]);
            byte[] bytes = built.ToArray();
            // Directory (sector 1) -> FAT sector (sector 0) -> back to directory: a 2-sector cycle in
            // the chain XlsCompoundFile.BuildWorkbook walks to read the OLE directory itself.
            PatchFatCycle(bytes, sectorA: 1, sectorB: 0);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => Excel.FromXls(new MemoryStream(bytes)));
            Assert.Contains("cycle", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        // SEC-8: [MS-CFB] fixes the mini sector size at 64 bytes (shift = 6, header offset 0x20).
        [Fact]
        public void MiniSectorShiftOtherThan64BytesThrows()
        {
            using MemoryStream built = XlsWorkbookBuilder.Build(sheets: [("S1", [["Alice", 1, true]])]);
            byte[] bytes = built.ToArray();
            BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(0x20), 7); // shift 7 -> 128-byte mini sectors

            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => Excel.FromXls(new MemoryStream(bytes)));
            Assert.Contains("sector size", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        // SEC-5: a run of zero-length CONTINUE records after an SST record never grows EnsureSharedCapacity's
        // buffer (needed <= buffer.Length stays true), so only the boundaries.Count-based charge added in
        // DecodeSstFromCursor can stop it. A tiny MaxSharedStringBytes makes this trip long before any real
        // memory pressure, without needing millions of records to prove the point.
        private static byte[] BuildFramedSstWithEmptyContinues(int continueCount)
        {
            using MemoryStream ms = new();
            void WriteRecord(int id, int payloadLength)
            {
                Span<byte> header = stackalloc byte[4];
                BinaryPrimitives.WriteUInt16LittleEndian(header, (ushort)id);
                BinaryPrimitives.WriteUInt16LittleEndian(header[2..], (ushort)payloadLength);
                ms.Write(header);
                if (payloadLength > 0)
                {
                    ms.Write(new byte[payloadLength]);
                }
            }
            WriteRecord(0x00FC, 8); // BIFF8 SST: cstTotal=0, cstUnique=0, no <si> data
            for (int i = 0; i < continueCount; i++)
            {
                WriteRecord(0x003C, 0); // zero-length CONTINUE
            }
            return ms.ToArray();
        }

        [Fact]
        public void ManyZeroLengthContinueRecordsTripSharedStringLimit()
        {
            byte[] framed = BuildFramedSstWithEmptyContinues(continueCount: 1000);
            using MemoryStream ms = XlsWorkbookBuilder.BuildRawSst(framed, labelSstCount: 0);
            var options = new ExcelReaderOptions { MaxSharedStringBytes = 100 };

            ExcelLimitExceededException ex = Assert.Throws<ExcelLimitExceededException>(
                () => Excel.FromXls(ms, options: options));
            Assert.Equal(nameof(ExcelReaderOptions.MaxSharedStringBytes), ex.LimitName);
        }

        [Fact]
        public void TooManyZipEntriesTripsMaxZipEntriesLimit()
        {
            using MemoryStream built = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            byte[] builtBytes = built.ToArray();
            using var ms = new MemoryStream();
            ms.Write(builtBytes, 0, builtBytes.Length);
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
            {
                for (int i = 0; i < 20; i++)
                {
                    using var entryStream = zip.CreateEntry($"junk/{i}.txt").Open();
                }
            }
            ms.Position = 0;

            var options = new ExcelReaderOptions { MaxZipEntries = 10 };

            ExcelLimitExceededException ex = Assert.Throws<ExcelLimitExceededException>(() =>
            {
                using XlsxReader reader = Excel.From(ms, options: options);
            });
            Assert.Equal(nameof(ExcelReaderOptions.MaxZipEntries), ex.LimitName);
            Assert.Equal(10, ex.Limit);
        }

        [Fact]
        public void CompressedSheetTripsTotalDecompressedLimit()
        {
            string value = new('A', 256 * 1024);
            using MemoryStream ms = WorkbookBuilder.Build(
                $"""<row r="1"><c r="A1" t="inlineStr"><is><t>{value}</t></is></c></row>""");
            Assert.True(ms.Length < 10_000);

            var options = new ExcelReaderOptions
            {
                MaxTotalDecompressedBytes = 16 * 1024,
            };

            using XlsxReader reader = Excel.From(ms, options: options);
            ExcelLimitExceededException ex = Assert.Throws<ExcelLimitExceededException>(() =>
            {
                using XlsxReader.Enumerator e = reader.GetEnumerator();
                Assert.True(e.MoveNext());
            });
            Assert.Equal(nameof(ExcelReaderOptions.MaxTotalDecompressedBytes), ex.LimitName);
        }

        [Fact]
        public void SharedStringsTripSharedStringLimit()
        {
            string value = new('B', 128 * 1024);
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v>0</v></c></row>""",
                sharedStrings: $"<si><t>{value}</t></si>");
            Assert.True(ms.Length < 10_000);

            var options = new ExcelReaderOptions
            {
                MaxTotalDecompressedBytes = 0,
                MaxSharedStringBytes = 1024,
            };

            using XlsxReader reader = Excel.From(ms, options: options);
            ExcelLimitExceededException ex = Assert.Throws<ExcelLimitExceededException>(() =>
            {
                using XlsxReader.Enumerator e = reader.GetEnumerator();
            });
            Assert.Equal(nameof(ExcelReaderOptions.MaxSharedStringBytes), ex.LimitName);
        }

        [Fact]
        public void ZeroTotalDecompressedLimitRestoresUnlimitedTotalRead()
        {
            string value = new('C', 96 * 1024);
            using MemoryStream ms = WorkbookBuilder.Build(
                $"""<row r="1"><c r="A1" t="inlineStr"><is><t>{value}</t></is></c></row>""");

            var options = new ExcelReaderOptions
            {
                MaxTotalDecompressedBytes = 0,
            };

            using XlsxReader reader = Excel.From(ms, options: options);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(value.Length, e.Current[0].GetString().Length);
        }

        [Fact]
        public async Task LimitedReadStreamCoversUnsupportedOperationsAndByteArrayAsyncRead()
        {
            await using MemoryStream inner = new([1, 2, 3, 4, 5]);
            await using LimitedReadStream limited = new(inner, new DecompressedByteCounter(limit: 10));

            Assert.True(limited.CanRead);
            Assert.False(limited.CanSeek);
            Assert.False(limited.CanWrite);
            Assert.Throws<NotSupportedException>(() => limited.Length);
            Assert.Throws<NotSupportedException>(() => limited.Position);
            Assert.Throws<NotSupportedException>(() => limited.Position = 0);
            Assert.Throws<NotSupportedException>(() => limited.Seek(0, SeekOrigin.Begin));
            Assert.Throws<NotSupportedException>(() => limited.SetLength(0));
            Assert.Throws<NotSupportedException>(() => limited.Write([1], 0, 1));

            byte[] buffer = new byte[3];
            int read = await limited.ReadAsync(buffer, 0, buffer.Length, TestContext.Current.CancellationToken);
            Assert.Equal(3, read);
            Assert.Equal([1, 2, 3], buffer);

            limited.Flush();
        }

        [Fact]
        public void LimitedReadStreamCountsSpanReadsAndEntryLimit()
        {
            using MemoryStream inner = new([1, 2, 3, 4]);
            using LimitedReadStream limited = new(
                inner,
                totalCounter: null,
                entryLimitName: nameof(ExcelReaderOptions.MaxCellBytes),
                entryLimit: 2);

            Span<byte> buffer = stackalloc byte[2];
            Assert.Equal(2, limited.Read(buffer));

            ExcelLimitExceededException ex;
            try
            {
                limited.Read(buffer);
                throw new Xunit.Sdk.XunitException("Expected ExcelLimitExceededException.");
            }
            catch (ExcelLimitExceededException caught)
            {
                ex = caught;
            }

            Assert.Equal(nameof(ExcelReaderOptions.MaxCellBytes), ex.LimitName);
            Assert.Equal(2, ex.Limit);
            Assert.Equal(4, ex.Actual);
        }

        // --- XlsCompoundFile (.xls / OLE-CFB) container-phase guard rails ---

        [Fact]
        public void ForgedWorkbookSizeNearUInt32MaxThrowsInvalidDataException()
        {
            using MemoryStream ms = XlsWorkbookBuilder.BuildPatched(
                XlsWorkbookBuilder.WorkbookSizeOffset, XlsWorkbookBuilder.LE64(0xFFFFFFFFL));
            Assert.Throws<InvalidDataException>(() => Excel.FromXls(ms));
        }

        [Fact]
        public void ForgedWorkbookSizeNearLongMaxThrowsInvalidDataException()
        {
            using MemoryStream ms = XlsWorkbookBuilder.BuildPatched(
                XlsWorkbookBuilder.WorkbookSizeOffset, XlsWorkbookBuilder.LE64(long.MaxValue - 1));
            Assert.Throws<InvalidDataException>(() => Excel.FromXls(ms));
        }

        [Fact]
        public void ForgedRootEntrySizeAboveIntMaxValueThrowsInvalidDataException()
        {
            // The root entry's mini-stream length used to be cast straight from a long to an int; a
            // value above the signed 32-bit range truncated to a negative limit, which the chain
            // reader interpreted as "read everything" instead of "read N bytes". The default builder's
            // workbook size always sits above the mini-stream cutoff, so the workbook size is shrunk
            // here too, forcing the one branch that reads the root entry's length at all.
            byte[] bytes = XlsWorkbookBuilder.Build(sheets: [("S1", [["A"]])]).ToArray();
            XlsWorkbookBuilder.LE64(100).CopyTo(bytes, XlsWorkbookBuilder.WorkbookSizeOffset);
            XlsWorkbookBuilder.LE64(int.MaxValue + 1L).CopyTo(bytes, XlsWorkbookBuilder.RootEntrySizeOffset);
            using var ms = new MemoryStream(bytes);
            Assert.Throws<InvalidDataException>(() => Excel.FromXls(ms));
        }

        [Fact]
        public void ForgedMiniCutoffThrowsInvalidDataException()
        {
            using MemoryStream ms = XlsWorkbookBuilder.BuildPatched(
                XlsWorkbookBuilder.MiniCutoffOffset, XlsWorkbookBuilder.LE32(int.MaxValue - 1));
            Assert.Throws<InvalidDataException>(() => Excel.FromXls(ms));
        }

        [Fact]
        public void ForgedOversizedWorkbookSizeTripsTotalDecompressedLimitBeforeAllocating()
        {
            using MemoryStream ms = XlsWorkbookBuilder.BuildPatched(
                XlsWorkbookBuilder.WorkbookSizeOffset, XlsWorkbookBuilder.LE64(5000));

            var options = new ExcelReaderOptions { MaxTotalDecompressedBytes = 4096 };

            ExcelLimitExceededException ex = Assert.Throws<ExcelLimitExceededException>(() => Excel.FromXls(ms, options: options));
            Assert.Equal(nameof(ExcelReaderOptions.MaxTotalDecompressedBytes), ex.LimitName);
            Assert.Equal(5000, ex.Actual);
        }

        [Fact]
        public void ExcelLimitExceededExceptionConstructorsAreCovered()
        {
            var empty = new ExcelLimitExceededException();
            Assert.Equal(string.Empty, empty.LimitName);

            var withMessage = new ExcelLimitExceededException("message");
            Assert.Equal("message", withMessage.Message);

            var inner = new InvalidOperationException("inner");
            var withInner = new ExcelLimitExceededException("outer", inner);
            Assert.Equal("outer", withInner.Message);
            Assert.Same(inner, withInner.InnerException);
        }
    }
}
