using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class MultiSheetTests
    {
        [Fact]
        public async Task SheetCountMatchesWorkbookSheetCount()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("Alpha", [[1]]),
                ("Beta", [[2]]));
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            Assert.Equal(2, reader.SheetCount);
        }

        [Fact]
        public async Task SheetNameMatchesCurrentSheet()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(("MySheet", []));
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            Assert.Equal("MySheet", reader.SheetName);
        }

        [Fact]
        public async Task TryMoveToSheetMatchesCaseInsensitively()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(("Sheet1", []));
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            Assert.True(reader.TryMoveToSheet("sheet1"));
            Assert.Equal("Sheet1", reader.SheetName);
        }

        [Fact]
        public async Task TryMoveToSheetReturnsFalseWhenNotFound()
        {
            await using var ms = await TypedWorkbook.BuildAsync();
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            Assert.False(reader.TryMoveToSheet("DoesNotExist"));
        }

        [Fact]
        public async Task MoveToSheetByIndexSwitchesCurrentSheet()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("First", []), ("Second", []));
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            reader.MoveToSheet(1);
            Assert.Equal("Second", reader.SheetName);
        }

        [Fact]
        public async Task MoveToSheetNegativeIndexThrows()
        {
            await using var ms = await TypedWorkbook.BuildAsync();
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.MoveToSheet(-1));
        }

        [Fact]
        public async Task MoveToSheetOutOfRangeIndexThrows()
        {
            await using var ms = await TypedWorkbook.BuildAsync();
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.MoveToSheet(1));
        }

        [Fact]
        public async Task SheetNavigationWorksThroughFormatAgnosticInterface()
        {
            // The whole point of #1: walk every sheet via IExcelRowReader (Excel.Open) with no downcast.
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("A", [[11]]),
                ("B", [[22]]));
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);

            var names = new List<string>();
            var firstCells = new List<int>();
            for (int i = 0; i < reader.SheetCount; i++)
            {
                reader.MoveToSheet(i);
                names.Add(reader.SheetName);
                await using var e = reader.GetEnumerator();
                Assert.True(await e.MoveNextAsync());
                Assert.True(e.Current[0].TryParse(null, out int v));
                firstCells.Add(v);
            }

            Assert.Equal(["A", "B"], names);
            Assert.Equal([11, 22], firstCells);
        }

        [Fact]
        public void CsvIsExposedAsSingleUnnamedSheet()
        {
            using var ms = new MemoryStream("h\nv\n"u8.ToArray());
            using IExcelRowReader reader = Excel.FromCsv(ms);

            Assert.Equal(1, reader.SheetCount);
            Assert.Equal("", reader.SheetName);
            reader.MoveToSheet(0);                                 // in range → no throw
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.MoveToSheet(1));
            Assert.True(reader.TryMoveToSheet(""));                // matches the one unnamed sheet
            Assert.False(reader.TryMoveToSheet("Sheet1"));         // no named sheets
        }

        [Fact]
        public async Task MultipleSheetsEachHaveDistinctData()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("A", [[11]]),
                ("B", [[22]]));
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            await using var e1 = reader.GetEnumerator();
            Assert.True(await e1.MoveNextAsync());
            Assert.True(e1.Current[0].TryParse(null, out int v1));
            Assert.Equal(11, v1);

            reader.MoveToSheet(1);

            await using var e2 = reader.GetEnumerator();
            Assert.True(await e2.MoveNextAsync());
            Assert.True(e2.Current[0].TryParse(null, out int v2));
            Assert.Equal(22, v2);
        }
    }
}
