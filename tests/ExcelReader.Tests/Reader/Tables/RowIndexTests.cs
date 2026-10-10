using System.Globalization;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using B = ExcelReader.Tests.Reader.Xlsb.Biff12Build;

namespace ExcelReader.Tests.Reader.Tables
{
    public class RowIndexTests
    {
        private sealed class StreamBackedMemory(byte[] bytes) : MemoryStream(bytes);

        [Fact]
        public void XlsxRowIndexFollowsTheRAttributeAndFallsBackToTheNextRow()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="2"><c r="A2"><v>1</v></c></row><row r="5"><c r="A5"><v>2</v></c></row><row><c r="A6"><v>3</v></c></row><row r="9"/>""");
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);

            Assert.Equal([1, 4, 5, 8], ReadIndexes(workbook.FirstSheet.GetEnumerator()));
        }

        [Fact]
        public void RowIndexStaysUnsetUntilEnabled()
        {
            using MemoryStream ms = WorkbookBuilder.Build("""<row r="3"><c r="A3"><v>1</v></c></row>""");
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            using XlsxWorkbook.Enumerator rows = workbook.FirstSheet.GetEnumerator();

            Assert.True(rows.MoveNext());
            Assert.Equal(-1, ((IRowIndexedEnumerator)rows).RowIndex);
        }

        [Fact]
        public void XlsbRowIndexFollowsTheRowHeaderAndFallsBackToTheNextRow()
        {
            using MemoryStream ms = XlsbSheet(
            [
                .. B.Record(Brt.RowHdr, B.U32(1)), .. B.Record(Brt.CellRk, B.CellRk(0, 0, (1u << 2) | 0x02)),
                .. B.Record(Brt.RowHdr, B.U32(4)), .. B.Record(Brt.CellRk, B.CellRk(0, 0, (2u << 2) | 0x02)),
                .. B.Record(Brt.RowHdr), .. B.Record(Brt.CellRk, B.CellRk(0, 0, (3u << 2) | 0x02)),
                .. B.Record(Brt.EndSheetData),
            ]);
            using XlsbWorkbook workbook = Excel.FromXlsb(ms);

            Assert.Equal([1, 4, 5], ReadIndexes(workbook.FirstSheet.GetEnumerator()));
        }

        [Fact]
        public async Task XlsxSyncAndAsyncRowIndexesMatchAcrossBufferRefills()
        {
            StringBuilder rows = new();
            for (int i = 0; i < 20_000; i++)
            {
                int r = (3 * i) + 1;
                rows.Append(CultureInfo.InvariantCulture, $"""<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>row-{i}-padding-padding-padding</t></is></c></row>""");
            }
            using MemoryStream source = WorkbookBuilder.Build(rows.ToString());
            byte[] bytes = source.ToArray();
            int[] expected = [.. Enumerable.Range(0, 20_000).Select(static i => 3 * i)];

            using XlsxWorkbook syncBook = Excel.FromXlsx(new StreamBackedMemory(bytes), leaveOpen: false);
            await using XlsxWorkbook asyncBook = await Excel.FromXlsxAsync(
                new StreamBackedMemory(bytes), leaveOpen: false, ct: TestContext.Current.CancellationToken);

            Assert.Equal(expected, ReadIndexes(syncBook.FirstSheet.GetEnumerator()));
            Assert.Equal(expected, await ReadIndexesAsync(asyncBook.FirstSheet.GetAsyncEnumerator(TestContext.Current.CancellationToken)));
        }

        [Fact]
        public async Task XlsbSyncAndAsyncRowIndexesMatchAcrossBufferRefills()
        {
            List<byte> sheet = [];
            for (int i = 0; i < 20_000; i++)
            {
                sheet.AddRange(B.Record(Brt.RowHdr, B.U32((uint)(3 * i))));
                sheet.AddRange(B.Record(Brt.CellSt, B.CellSt(0, 0, $"row-{i}-padding-padding-padding")));
            }
            sheet.AddRange(B.Record(Brt.EndSheetData));
            byte[] bytes = XlsbSheet([.. sheet]).ToArray();
            int[] expected = [.. Enumerable.Range(0, 20_000).Select(static i => 3 * i)];

            using XlsbWorkbook syncBook = Excel.FromXlsb(new StreamBackedMemory(bytes), leaveOpen: false);
            await using XlsbWorkbook asyncBook = await Excel.FromXlsbAsync(
                new StreamBackedMemory(bytes), leaveOpen: false, ct: TestContext.Current.CancellationToken);

            Assert.Equal(expected, ReadIndexes(syncBook.FirstSheet.GetEnumerator()));
            Assert.Equal(expected, await ReadIndexesAsync(asyncBook.FirstSheet.GetAsyncEnumerator(TestContext.Current.CancellationToken)));
        }

        [Fact]
        public void FixtureRowIndexesSkipAbsentRows()
        {
            using IExcelWorkbook workbook = Excel.Open(TableMetadataTests.FixturePath("tables.xlsx"));

            Assert.Equal([3, 4, 5, 6, 7, 10], ReadIndexes(workbook.FirstSheet.GetEnumerator()));
        }

        private static int[] ReadIndexes(IExcelRowEnumerator rows)
        {
            using (rows)
            {
                IRowIndexedEnumerator indexed = (IRowIndexedEnumerator)rows;
                indexed.EnableRowIndex();
                List<int> indexes = [];
                while (rows.MoveNext())
                {
                    indexes.Add(indexed.RowIndex);
                }
                return [.. indexes];
            }
        }

        private static async Task<int[]> ReadIndexesAsync(IExcelRowEnumerator rows)
        {
            await using (rows.ConfigureAwait(false))
            {
                IRowIndexedEnumerator indexed = (IRowIndexedEnumerator)rows;
                indexed.EnableRowIndex();
                List<int> indexes = [];
                while (await rows.MoveNextAsync().ConfigureAwait(false))
                {
                    indexes.Add(indexed.RowIndex);
                }
                return [.. indexes];
            }
        }

        private static MemoryStream XlsbSheet(byte[] sheet)
        {
            MemoryStream ms = new();
            using (var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            {
                Add(zip, "xl/workbook.bin", B.Record(Brt.BundleSh, [.. B.U32(0), .. B.U32(0), .. B.WideString("rId1"), .. B.WideString("S1")]));
                Add(zip, "xl/_rels/workbook.bin.rels", Encoding.UTF8.GetBytes(
                    """<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.bin"/></Relationships>"""));
                Add(zip, "xl/styles.bin", []);
                Add(zip, "xl/sharedStrings.bin", []);
                Add(zip, "xl/worksheets/sheet1.bin", sheet);
            }
            ms.Position = 0;
            return ms;
        }

        private static void Add(System.IO.Compression.ZipArchive zip, string name, byte[] bytes)
        {
            using Stream stream = zip.CreateEntry(name).Open();
            stream.Write(bytes);
        }
    }
}
