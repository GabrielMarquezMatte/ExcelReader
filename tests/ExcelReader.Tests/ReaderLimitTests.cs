using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Crypto;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    public class ReaderLimitTests
    {
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
        [Theory]
        [InlineData("AAAA1")]
        [InlineData("MWLRALP1")]
        public void FourOrMoreColumnLettersThrowColumnLimit(string cellRef)
        {
            using MemoryStream built = WorkbookBuilder.Build($"""<row r="1"><c r="{cellRef}"><v>1</v></c></row>""");

            using XlsxReader reader = Excel.From(built);
            Assert.Throws<ExcelLimitExceededException>(() =>
            {
                using XlsxReader.Enumerator e = reader.GetEnumerator();
                Assert.True(e.MoveNext());
            });
        }

        [Fact]
        public void LastColumnXfdStillReads()
        {
            using MemoryStream built = WorkbookBuilder.Build("""<row r="1"><c r="XFD1"><v>1</v></c></row>""");

            using XlsxReader reader = Excel.From(built);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(16_384, e.Current.ColumnCount);
        }

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

        private static void PatchFatCycle(byte[] oleBytes, int sectorA, int sectorB)
        {
            const int SectorSize = 512;
            const int FatSectorOffset = SectorSize;
            BinaryPrimitives.WriteInt32LittleEndian(oleBytes.AsSpan(FatSectorOffset + (sectorA * 4)), sectorB);
            BinaryPrimitives.WriteInt32LittleEndian(oleBytes.AsSpan(FatSectorOffset + (sectorB * 4)), sectorA);
        }

        [Fact]
        public void FatChainCycleThrowsInsteadOfUnboundedAllocation()
        {
            using MemoryStream built = XlsWorkbookBuilder.Build(sheets: [("S1", [["Alice", 1, true]])]);
            byte[] bytes = built.ToArray();
            PatchFatCycle(bytes, sectorA: 1, sectorB: 0);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => Excel.FromXls(new MemoryStream(bytes)));
            Assert.Contains("cycle", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void MiniSectorShiftOtherThan64BytesThrows()
        {
            using MemoryStream built = XlsWorkbookBuilder.Build(sheets: [("S1", [["Alice", 1, true]])]);
            byte[] bytes = built.ToArray();
            BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(0x20), 7);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(
                () => Excel.FromXls(new MemoryStream(bytes)));
            Assert.Contains("sector size", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

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
            WriteRecord(0x00FC, 8);
            for (int i = 0; i < continueCount; i++)
            {
                WriteRecord(0x003C, 0);
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


        [Fact]
        public void Should_Throw_When_Encrypted_Package_Declares_More_Plaintext_Than_Budget()
        {
            byte[] bytes = EncryptedFixtures.Bytes("agile-aes256-sha512.xlsx");
            var tiny = ExcelReaderOptions.Default with
            {
                Password = EncryptedFixtures.Password,
                MaxTotalDecompressedBytes = 512,
            };
            Assert.Throws<ExcelLimitExceededException>(() => Excel.Open(bytes, tiny));
        }

        [Fact]
        public void Should_Throw_When_SpinCount_Above_Configured_Cap()
        {
            byte[] bytes = EncryptedFixtures.Bytes("agile-aes256-sha512.xlsx");
            var tight = ExcelReaderOptions.Default with
            {
                Password = EncryptedFixtures.Password,
                MaxPasswordSpinCount = 1,
            };
            Assert.Throws<ExcelLimitExceededException>(() => Excel.Open(bytes, tight));
        }

        [Fact]
        public void Should_Ignore_SpinCount_Cap_When_Scheme_Is_Standard()
        {
            byte[] container = StandardContainerFixture.Forge();
            ExcelReaderOptions options = new()
            {
                Password = StandardContainerFixture.Password,
                MaxPasswordSpinCount = 1,
            };

            using IExcelRowReader reader = Excel.Open(container, options);

            Assert.True(reader.SheetCount > 0);
        }

        public static TheoryData<string> EncryptedMutationFixtures()
        {
            var data = new TheoryData<string>();
            foreach (string name in EncryptedFixtures.All)
            {
                data.Add(name);
            }
            return data;
        }

        [Theory]
        [MemberData(nameof(EncryptedMutationFixtures))]
        [SuppressMessage("Design", "CA1031:Do not catch general exception types",
            Justification = "Classifying every escaping exception as acceptable-rejection or bug IS this " +
                "test's oracle; the unacceptable ones are recorded and rethrown after the loop rather than " +
                "inside the handler, so the reported round stays deterministic.")]
        public void Should_Reject_Gracefully_When_Encrypted_Container_Mutated(string fixture)
        {
            const int Rounds = 340;
            byte[] seed = EncryptedFixtures.Bytes(fixture);
            var options = ExcelReaderOptions.Default with { Password = EncryptedFixtures.PasswordFor(fixture) };
            var rng = new Random(20260826);
            byte[][] mutations = new byte[Rounds][];
            int[][] mutatedOffsets = new int[Rounds][];
            for (int i = 0; i < Rounds; i++)
            {
                mutations[i] = FuzzMutation.MutateCopy(seed, rng, out int[] positions);
                mutatedOffsets[i] = positions;
            }
            Exception?[] failures = new Exception?[Rounds];
            int completed = 0;
            ParallelOptions parallelOptions = new()
            {
                MaxDegreeOfParallelism = Math.Max(2, Environment.ProcessorCount / 2),
                CancellationToken = TestContext.Current.CancellationToken,
            };
            Parallel.For(0, Rounds, parallelOptions, i =>
            {
                try
                {
                    using IExcelRowReader reader = Excel.Open(mutations[i], options);
                    foreach (Row row in reader)
                    {
                        foreach (RowCell cell in row.Cells)
                        {
                            _ = cell.Value.GetString();
                        }
                    }
                }
                catch (Exception ex) when (FuzzMutation.IsAcceptable(ex))
                {
                }
                catch (Exception ex)
                {
                    failures[i] = ex;
                    return;
                }
                Interlocked.Increment(ref completed);
            });

            int firstFailure = Array.FindIndex(failures, ex => ex is not null);
            if (firstFailure >= 0)
            {
                Exception failure = failures[firstFailure]!;
                string offsets = string.Join(", ", mutatedOffsets[firstFailure]);
                throw new InvalidOperationException(
                    string.Create(CultureInfo.InvariantCulture,
                        $"Round {firstFailure} on fixture '{fixture}' produced an unacceptable '{failure.GetType().Name}' (mutated byte offsets: [{offsets}]): {failure.Message}"),
                    failure);
            }
            Assert.Equal(Rounds, completed);
        }

        [Fact]
        public void MiniStreamSectorNearIntMaxThrowsInvalidDataInsteadOfOverflow()
        {
            byte[] miniStream = new byte[128];
            int[] miniFat = [-1];
            const int HugeSector = 40_000_000;
            Assert.Throws<InvalidDataException>(() =>
                CfbContainer.ReadMiniStream(miniStream, miniFat, miniSectorSize: 64, startSector: HugeSector, size: 64));
        }

        [Fact]
        public void MiniStreamSectorNearEndOfStreamThrowsInvalidDataInsteadOfRangeError()
        {
            byte[] miniStream = new byte[100];
            int[] miniFat = [-1];
            Assert.Throws<InvalidDataException>(() =>
                CfbContainer.ReadMiniStream(miniStream, miniFat, miniSectorSize: 64, startSector: 1, size: 64));
        }
    }
}
