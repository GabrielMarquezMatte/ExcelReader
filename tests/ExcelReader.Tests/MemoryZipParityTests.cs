using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    public class MemoryZipParityTests
    {
        public static IEnumerable<object[]> Fixtures
        {
            get
            {
                yield return [new MemoryFixture("xlsx", BuildXlsx, OpenXlsxStream, OpenXlsxMemory)];
                yield return [new MemoryFixture("xlsb", BuildXlsb, OpenXlsbStream, OpenXlsbMemory)];
            }
        }


        [Theory]
        [MemberData(nameof(Fixtures))]
        public void MemoryReadMatchesStreamedRead(MemoryFixture fixture)
        {
            byte[] bytes = fixture.Build();

            List<CellSnapshot> streamed = ReadViaStream(bytes, fixture.OpenStream);
            List<CellSnapshot> memory = ReadViaMemory(bytes, fixture.OpenMemory);

            Assert.NotEmpty(streamed);
            Assert.Equal(streamed, memory);
        }

        [Theory]
        [MemberData(nameof(Fixtures))]
        public async Task MemoryAsyncReadMatchesStreamedRead(MemoryFixture fixture)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = fixture.Build();

            List<CellSnapshot> streamed = ReadViaStream(bytes, fixture.OpenStream);
            List<CellSnapshot> memory = await ReadViaMemoryAsync(bytes, fixture.OpenMemory, ct);

            Assert.Equal(streamed, memory);
        }

        [Theory]
        [MemberData(nameof(Fixtures))]
        public async Task MemoryAsyncEnumeratorNeverSuspends(MemoryFixture fixture)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = fixture.Build();
            using IExcelRowReader reader = fixture.OpenMemory(bytes, ExcelReaderOptions.Default);
            ValueTask<IExcelRowEnumerator> task = reader.GetAsyncEnumeratorAsync(ct);
            Assert.True(task.IsCompleted);
            await using IExcelRowEnumerator e = await task;
        }


        [Theory]
        [MemberData(nameof(Fixtures))]
        public void PrefetchDecompressionProducesIdenticalOutput(MemoryFixture fixture)
        {
            byte[] bytes = fixture.Build();
            var options = new ExcelReaderOptions { PrefetchDecompression = true };

            List<CellSnapshot> off = ReadViaMemory(bytes, fixture.OpenMemory, ExcelReaderOptions.Default);
            List<CellSnapshot> on = ReadViaMemory(bytes, fixture.OpenMemory, options);

            Assert.Equal(off, on);
        }


        [Fact]
        public void MoveToSheetAndTryMoveToSheetWorkOnTheMemoryPath()
        {
            using MemoryStream built = WorkbookBuilder.BuildMultiSheet(
                sheets:
                [
                    ("First", """<row r="1"><c r="A1"><v>1</v></c></row>"""),
                    ("Second", """<row r="1"><c r="A1"><v>2</v></c></row>"""),
                ]);
            byte[] bytes = built.ToArray();

            using XlsxReader reader = Excel.FromXlsx(bytes.AsMemory());
            Assert.Equal(2, reader.SheetCount);

            reader.MoveToSheet(1);
            Assert.Equal("Second", reader.SheetName);
            using (XlsxReader.Enumerator e = reader.GetEnumerator())
            {
                Assert.True(e.MoveNext());
                Assert.Equal("2", e.Current[0].GetString());
            }

            Assert.True(reader.TryMoveToSheet("First"));
            Assert.Equal("First", reader.SheetName);
            using (XlsxReader.Enumerator e = reader.GetEnumerator())
            {
                Assert.True(e.MoveNext());
                Assert.Equal("1", e.Current[0].GetString());
            }

            Assert.False(reader.TryMoveToSheet("NoSuchSheet"));
        }


        [Fact]
        public void ReadsCorrectlyOverANonArrayBackedMemory()
        {
            using MemoryStream built = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>42</v></c></row>""");
            byte[] bytes = built.ToArray();
            var manager = new NonArrayMemoryManager(bytes);

            using XlsxReader reader = Excel.FromXlsx(manager.Memory);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("42", e.Current[0].GetString());
        }


        [Fact]
        public void OpenMemoryDetectsXlsx()
        {
            using MemoryStream built = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            byte[] bytes = built.ToArray();
            using IExcelRowReader reader = Excel.Open(bytes.AsMemory());
            Assert.IsType<XlsxReader>(reader);
        }

        [Fact]
        public async Task OpenMemoryDetectsXlsb()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = await BuildXlsbAsync(ct);
            using IExcelRowReader reader = Excel.Open(bytes.AsMemory());
            Assert.IsType<XlsbReader>(reader);
        }

        [Fact]
        public void OpenMemoryDetectsXls()
        {
            using MemoryStream built = XlsWorkbookBuilder.Build(sheets: [("S1", [["Name", 1, true]])]);
            byte[] bytes = built.ToArray();
            using IExcelRowReader reader = Excel.Open(bytes.AsMemory());
            Assert.IsType<XlsReader>(reader);
        }

        [Fact]
        public void OpenMemoryThrowsInvalidDataExceptionForUnknownSignature()
        {
            byte[] bytes = "not a workbook"u8.ToArray();
            Assert.Throws<InvalidDataException>(() => Excel.Open(bytes.AsMemory()));
        }

        [Theory]
        [MemberData(nameof(Fixtures))]
        public void OpenMemoryMatchesOpenStreamForAutoDetection(MemoryFixture fixture)
        {
            byte[] bytes = fixture.Build();
            using IExcelRowReader streamed = Excel.Open(new MemoryStream(bytes, writable: false));
            using IExcelRowReader memory = Excel.Open(bytes.AsMemory());
            Assert.Equal(streamed.GetType(), memory.GetType());
        }


        [Theory]
        [MemberData(nameof(Fixtures))]
        public void DoubleDisposeIsSafe(MemoryFixture fixture)
        {
            byte[] bytes = fixture.Build();
            IExcelRowReader reader = fixture.OpenMemory(bytes, ExcelReaderOptions.Default);
            reader.Dispose();
            Assert.Null(Record.Exception(reader.Dispose));
        }

        [Theory]
        [MemberData(nameof(Fixtures))]
        public void DoubleDisposingTheEnumeratorIsSafe(MemoryFixture fixture)
        {
            byte[] bytes = fixture.Build();
            using IExcelRowReader reader = fixture.OpenMemory(bytes, ExcelReaderOptions.Default);
            using IExcelRowEnumerator e = reader.GetEnumerator();
            e.Dispose();
            Assert.Null(Record.Exception(e.Dispose));
        }


        [Fact]
        public void ForgedOversizedEntryThrowsSameLimitOnBothPaths()
        {
            using MemoryStream built = WorkbookBuilder.Build(
                """<row r="1"><c r="A1"><v>1</v></c></row>""",
                styles: """<styleSheet><cellXfs count="1"><xf/></cellXfs></styleSheet>""");
            byte[] bytes = built.ToArray();
            ForgeCentralDirectoryUncompressedSize(bytes, "xl/styles.xml", 50_000_000);
            var options = new ExcelReaderOptions { MaxTotalDecompressedBytes = 4096 };

            ExcelLimitExceededException streamedEx = Assert.Throws<ExcelLimitExceededException>(() =>
            {
                using XlsxReader reader = Excel.FromXlsx(new MemoryStream(bytes), options: options);
            });
            ExcelLimitExceededException memoryEx = Assert.Throws<ExcelLimitExceededException>(() =>
            {
                using XlsxReader reader = Excel.FromXlsx(bytes.AsMemory(), options);
            });
            Assert.Equal(streamedEx.LimitName, memoryEx.LimitName);
        }

        [Fact]
        public void MissingWorksheetPartThrowsInvalidDataExceptionOnBothPaths()
        {
            using MemoryStream built = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            byte[] bytes = built.ToArray();
            RemoveZipEntry(bytes, "xl/worksheets/sheet1.xml", out byte[] withoutSheet);

            Assert.Throws<InvalidDataException>(() =>
            {
                using XlsxReader reader = Excel.FromXlsx(new MemoryStream(withoutSheet));
                using XlsxReader.Enumerator e = reader.GetEnumerator();
            });
            Assert.Throws<InvalidDataException>(() =>
            {
                using XlsxReader reader = Excel.FromXlsx(withoutSheet.AsMemory());
                using XlsxReader.Enumerator e = reader.GetEnumerator();
            });
        }


        [Fact]
        public void MutatedZipBytesNeverCrashExcelOpenMemory()
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
                    FuzzMutation.RunBounded(() => OpenAndDrainMemory(mutated));
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

        private static void OpenAndDrainMemory(byte[] bytes)
        {
            using IExcelRowReader reader = Excel.Open(bytes.AsMemory());
            for (int s = 0; s < reader.SheetCount; s++)
            {
                reader.MoveToSheet(s);
                using IExcelRowEnumerator e = reader.GetEnumerator();
                while (e.MoveNext())
                {
                    for (int c = 0; c < e.Current.ColumnCount; c++)
                    {
                        _ = e.Current[c].GetString();
                    }
                }
            }
        }


        [Fact]
        public void RealSampleWorkbookMatchesBetweenStreamAndMemory()
        {
            string path = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");
            byte[] bytes = File.ReadAllBytes(path);

            List<CellSnapshot> streamed = ReadViaStream(bytes, OpenXlsxStream);
            List<CellSnapshot> memory = ReadViaMemory(bytes, OpenXlsxMemory);

            Assert.NotEmpty(streamed);
            Assert.Equal(streamed, memory);
        }


        private static List<CellSnapshot> ReadViaStream(byte[] bytes, Func<Stream, ExcelReaderOptions, IExcelRowReader> open)
        {
            using MemoryStream stream = new(bytes, writable: false);
            using IExcelRowReader reader = open(stream, ExcelReaderOptions.Default);
            using IExcelRowEnumerator e = reader.GetEnumerator();
            List<CellSnapshot> cells = [];
            int rowIndex = 0;
            while (e.MoveNext())
            {
                AddRow(cells, rowIndex++, e.Current);
            }
            return cells;
        }

        private static List<CellSnapshot> ReadViaMemory(
            byte[] bytes, Func<byte[], ExcelReaderOptions, IExcelRowReader> open, ExcelReaderOptions? options = null)
        {
            using IExcelRowReader reader = open(bytes, options ?? ExcelReaderOptions.Default);
            using IExcelRowEnumerator e = reader.GetEnumerator();
            List<CellSnapshot> cells = [];
            int rowIndex = 0;
            while (e.MoveNext())
            {
                AddRow(cells, rowIndex++, e.Current);
            }
            return cells;
        }

        private static async Task<List<CellSnapshot>> ReadViaMemoryAsync(
            byte[] bytes, Func<byte[], ExcelReaderOptions, IExcelRowReader> open, CancellationToken ct)
        {
            using IExcelRowReader reader = open(bytes, ExcelReaderOptions.Default);
            await using IExcelRowEnumerator e = await reader.GetAsyncEnumeratorAsync(ct);
            List<CellSnapshot> cells = [];
            int rowIndex = 0;
            while (await e.MoveNextAsync())
            {
                AddRow(cells, rowIndex++, e.Current);
            }
            return cells;
        }

        private static void AddRow(List<CellSnapshot> cells, int rowIndex, Row row)
        {
            cells.Add(CellSnapshot.RowMarker(rowIndex, row.ColumnCount));
            for (int column = 0; column < row.ColumnCount; column++)
            {
                Cell cell = row[column];
                bool hasDouble = cell.TryGetDouble(out double value);
                cells.Add(new CellSnapshot(
                    rowIndex,
                    column,
                    row.ColumnCount,
                    cell.Type,
                    cell.GetString(),
                    hasDouble,
                    hasDouble ? BitConverter.DoubleToInt64Bits(value) : 0));
            }
        }


        private static byte[] BuildXlsx()
        {
            using MemoryStream built = WorkbookBuilder.BuildMultiSheet(
                sheets:
                [
                    ("S1", """<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c></row>"""),
                    ("S2", """<row r="1"><c r="A1"><v>7</v></c></row>"""),
                ],
                sharedStrings: "<si><t>hello</t></si><si><t>world</t></si>",
                styles: "<styleSheet><cellXfs count=\"1\"><xf numFmtId=\"14\"/></cellXfs></styleSheet>");
            return built.ToArray();
        }

        private static byte[] BuildXlsb()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            return BuildXlsbAsync(ct).GetAwaiter().GetResult();
        }

        private static async Task<byte[]> BuildXlsbAsync(CancellationToken ct)
        {
            MemoryStream ms = new();
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsbSheetWriter sheet = wb.AddSheet("Sheet1");
                await sheet.StartAsync(ct);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(ct))
                {
                    row.Write("hello");
                    row.Write(42);
                    row.Write(true);
                    row.Write(new DateTime(2026, 1, 1));
                }
                await sheet.EndAsync(ct);
                await wb.EndAsync(ct);
            }
            return ms.ToArray();
        }


        private static IExcelRowReader OpenXlsxStream(Stream stream, ExcelReaderOptions options)
        {
            return Excel.FromXlsx(stream, options: options);
        }

        private static IExcelRowReader OpenXlsxMemory(byte[] bytes, ExcelReaderOptions options)
        {
            return Excel.FromXlsx(bytes.AsMemory(), options);
        }

        [SuppressMessage("Performance", "CA1859:Use concrete types when possible for improved performance",
            Justification = "Must match the Func<Stream, ExcelReaderOptions, IExcelRowReader> delegate shape shared with the XLSX fixture.")]
        private static IExcelRowReader OpenXlsbStream(Stream stream, ExcelReaderOptions options)
        {
            return Excel.FromXlsb(stream, options: options);
        }

        [SuppressMessage("Performance", "CA1859:Use concrete types when possible for improved performance",
            Justification = "Must match the Func<byte[], ExcelReaderOptions, IExcelRowReader> delegate shape shared with the XLSX fixture.")]
        private static IExcelRowReader OpenXlsbMemory(byte[] bytes, ExcelReaderOptions options)
        {
            return Excel.FromXlsb(bytes.AsMemory(), options);
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

        private static void RemoveZipEntry(byte[] zipBytes, string entryName, out byte[] result)
        {
            using var ms = new MemoryStream();
            ms.Write(zipBytes, 0, zipBytes.Length);
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
            {
                zip.GetEntry(entryName)?.Delete();
            }
            result = ms.ToArray();
        }

        public sealed record MemoryFixture(
            string Name,
            Func<byte[]> Build,
            Func<Stream, ExcelReaderOptions, IExcelRowReader> OpenStream,
            Func<byte[], ExcelReaderOptions, IExcelRowReader> OpenMemory)
        {
            public override string ToString()
            {
                return Name;
            }
        }

    }
}
