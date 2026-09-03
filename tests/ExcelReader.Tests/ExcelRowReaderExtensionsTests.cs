using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class ExcelRowReaderExtensionsTests
    {
        [Fact]
        public async Task SheetsYieldsEveryIndexAndNameInOrder()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("Alpha", [[1]]),
                ("Beta", [[2]]),
                ("Gamma", [[3]]));
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);

            List<ExcelSheet> sheets = [.. reader.Sheets()];

            Assert.Equal(
                [new ExcelSheet(0, "Alpha"), new ExcelSheet(1, "Beta"), new ExcelSheet(2, "Gamma")],
                sheets);
        }

        [Fact]
        public async Task SheetsSelectsEachSheetBeforeYieldingItSoRowsAreReadable()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("A", [[11]]),
                ("B", [[22]]));
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);

            var firstCells = new List<int>();
            foreach (ExcelSheet sheet in reader.Sheets())
            {
                await using var e = reader.GetEnumerator();
                Assert.True(await e.MoveNextAsync());
                Assert.True(e.Current[0].TryParse(null, out int v));
                firstCells.Add(v);
            }

            Assert.Equal([11, 22], firstCells);
        }

        [Fact]
        public async Task SheetsOnASingleSheetWorkbookYieldsOneEntry()
        {
            await using var ms = await TypedWorkbook.BuildAsync([1]);
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);

            ExcelSheet[] sheets = [.. reader.Sheets()];

            Assert.Single(sheets);
            Assert.Equal(0, sheets[0].Index);
        }

        [Fact]
        public void SheetsThrowsOnNullReader()
        {
            IExcelRowReader reader = null!;
            Assert.Throws<ArgumentNullException>(() => reader.Sheets());
        }
    }
}
