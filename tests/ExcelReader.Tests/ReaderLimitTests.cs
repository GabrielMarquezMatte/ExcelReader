using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

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
                    () => acc.Add(16_384, start: 0, len: 0, CellType.ExcelString, style: 0, fromShared: false));
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
                acc.Add(16_383, start: 0, len: 0, CellType.ExcelString, style: 0, fromShared: false);
                Assert.Equal(1, acc.Count);
            }
            finally
            {
                acc.Return();
            }
        }
        // Patches the uncompressed-size field of a central-directory record in place, so the entry's
        // declared size (what ZipArchiveEntry.Length reports) lies far above its real, tiny compressed
        // content — the exact shape of a zip-bomb-style amplification attack (see docs/road-to-a.md, F1).
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
