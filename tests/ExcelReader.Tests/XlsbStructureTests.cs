using System.Buffers.Binary;
using System.IO.Compression;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    /// <summary>
    /// Checks the sheet records strict readers validate: BrtWsDim present before the sheet data, and
    /// every BrtRowHdr laid out per MS-XLSB with column spans that cover its cells.
    /// </summary>
    public class XlsbStructureTests
    {
        private sealed record SheetStructure(uint[] Dimension, List<RowStructure> Rows);

        private sealed record RowStructure(int Row, List<(int ColMic, int ColLast)> Spans, List<int> CellColumns);

        [Fact]
        public void ExcelWrittenWorkbookPassesTheChecks()
        {
            SheetStructure sheet = ReadStructure(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "data", "RealExcel.xlsb")));

            Assert.Equal([0u, 100u, 0u, 17u], sheet.Dimension);
            AssertRowsCovered(sheet);
        }

        [Fact]
        public async Task SmallSheetGetsItsExactUsedRange()
        {
            byte[] bytes = await WriteAsync(sheet =>
            {
                for (int r = 0; r < 10; r++)
                {
                    using XlsbRowWriter row = sheet.StartRow();
                    row.Write("name");
                    row.Write(r);
                    row.Write(r * 1.5);
                }
            });

            SheetStructure structure = ReadStructure(bytes);

            Assert.Equal([0u, 9u, 0u, 2u], structure.Dimension);
            Assert.All(structure.Rows, row => Assert.Equal([(0, 2)], row.Spans));
            AssertRowsCovered(structure);
        }

        [Fact]
        public async Task SkippedLeadingColumnsNarrowTheSpanAndRange()
        {
            byte[] bytes = await WriteAsync(sheet =>
            {
                using XlsbRowWriter row = sheet.StartRow();
                row.Skip(3);
                row.Write("d");
                row.Write("e");
            });

            SheetStructure structure = ReadStructure(bytes);

            Assert.Equal([0u, 0u, 3u, 4u], structure.Dimension);
            Assert.Equal([(3, 4)], structure.Rows.Single().Spans);
        }

        [Fact]
        public async Task EmptySheetReportsA1()
        {
            byte[] bytes = await WriteAsync(_ => { });

            Assert.Equal([0u, 0u, 0u, 0u], ReadStructure(bytes).Dimension);
        }

        [Fact]
        public async Task RowCrossingColumnBlocksGetsOneSpanPerBlock()
        {
            byte[] bytes = await WriteAsync(sheet =>
            {
                using (XlsbRowWriter row = sheet.StartRow())
                {
                    row.Write("a");
                    row.Skip(2046);
                    row.Write("x");
                }
                using (XlsbRowWriter row = sheet.StartRow())
                {
                    row.Write("after");
                }
            });

            SheetStructure structure = ReadStructure(bytes);

            Assert.Equal([(0, 1023), (1024, 2047)], structure.Rows[0].Spans);
            Assert.Equal([(0, 0)], structure.Rows[1].Spans);
            AssertRowsCovered(structure);
            using XlsbReader reader = Excel.FromXlsb(new MemoryStream(bytes));
            var values = new List<string>();
            foreach (var r in reader)
            {
                values.Add(r[0].GetString());
            }
            Assert.Equal(["a", "after"], values);
        }

        [Fact]
        public async Task SpilledSheetLeavesTheRowRangeOpen()
        {
            byte[] bytes = await WriteAsync(sheet =>
            {
                for (int r = 0; r < 20_000; r++)
                {
                    using XlsbRowWriter row = sheet.StartRow();
                    row.Write("row text");
                    row.Write(r);
                }
            });

            SheetStructure structure = ReadStructure(bytes);

            Assert.Equal([0u, 1_048_575u, 0u, 1u], structure.Dimension);
            Assert.Equal(20_000, structure.Rows.Count);
            AssertRowsCovered(structure);
        }

        private static async Task<byte[]> WriteAsync(Action<XlsbSheetWriter> build)
        {
            using MemoryStream ms = new();
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                build(sheet);
                await sheet.EndAsync(TestContext.Current.CancellationToken);
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }
            return ms.ToArray();
        }

        private static void AssertRowsCovered(SheetStructure sheet)
        {
            foreach (RowStructure row in sheet.Rows)
            {
                Assert.All(row.Spans, span =>
                {
                    Assert.True(span.ColMic <= span.ColLast, $"row {row.Row}: span {span} is inverted");
                    Assert.True(span.ColMic >> 10 == span.ColLast >> 10, $"row {row.Row}: span {span} crosses a 1024-column block");
                });
                Assert.All(row.CellColumns, col =>
                    Assert.True(row.Spans.Any(s => s.ColMic <= col && col <= s.ColLast), $"row {row.Row}: column {col} is outside its spans"));
            }
        }

        private static SheetStructure ReadStructure(byte[] workbook)
        {
            using var zip = new ZipArchive(new MemoryStream(workbook), ZipArchiveMode.Read);
            ZipArchiveEntry entry = zip.Entries.Single(e => e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal));
            using var sheetBytes = new MemoryStream();
            using (Stream s = entry.Open())
            {
                s.CopyTo(sheetBytes);
            }
            return Parse(sheetBytes.ToArray());
        }

        private static SheetStructure Parse(ReadOnlySpan<byte> sheet)
        {
            uint[]? dimension = null;
            bool sawBeginSheet = false;
            bool inData = false;
            var rows = new List<RowStructure>();
            var reader = new Biff12RecordReader(sheet);
            while (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
            {
                switch (id)
                {
                    case Brt.BeginSheet:
                        sawBeginSheet = true;
                        break;
                    case Brt.WsDim:
                        Assert.True(sawBeginSheet && !inData, "BrtWsDim must sit between BrtBeginSheet and BrtBeginSheetData");
                        Assert.Equal(16, payload.Length);
                        dimension = new uint[4];
                        for (int i = 0; i < 4; i++)
                        {
                            dimension[i] = BinaryPrimitives.ReadUInt32LittleEndian(payload.Slice(i * 4, 4));
                        }
                        break;
                    case Brt.BeginSheetData:
                        Assert.NotNull(dimension);
                        inData = true;
                        break;
                    case Brt.RowHdr:
                        int spanCount = (int)BinaryPrimitives.ReadUInt32LittleEndian(payload.Slice(13, 4));
                        Assert.Equal(17 + (spanCount * 8), payload.Length);
                        var spans = new List<(int, int)>();
                        for (int i = 0; i < spanCount; i++)
                        {
                            ReadOnlySpan<byte> span = payload.Slice(17 + (i * 8), 8);
                            spans.Add(((int)BinaryPrimitives.ReadUInt32LittleEndian(span), (int)BinaryPrimitives.ReadUInt32LittleEndian(span[4..])));
                        }
                        rows.Add(new RowStructure((int)BinaryPrimitives.ReadUInt32LittleEndian(payload), spans, []));
                        break;
                    case >= Brt.CellBlank and <= Brt.FmlaError:
                        rows[^1].CellColumns.Add((int)BinaryPrimitives.ReadUInt32LittleEndian(payload));
                        break;
                }
            }
            Assert.NotNull(dimension);
            return new SheetStructure(dimension, rows);
        }
    }
}
