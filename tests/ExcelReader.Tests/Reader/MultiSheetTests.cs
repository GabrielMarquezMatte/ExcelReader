using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Tests.Reader
{
    public class MultiSheetTests
    {
        [Fact]
        public async Task SheetCountMatchesWorkbookSheetCount()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("Alpha", [[1]]),
                ("Beta", [[2]]));
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            Assert.Equal(2, reader.SheetCount);
        }

        [Fact]
        public async Task FirstSheetCarriesTheWorkbookSheetName()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(("MySheet", []));
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            Assert.Equal("MySheet", reader.FirstSheet.Name);
        }

        [Fact]
        public async Task TryGetSheetMatchesCaseInsensitively()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(("Sheet1", []));
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            Assert.True(reader.TryGetSheet("sheet1", out XlsxSheet sheet));
            Assert.Equal("Sheet1", sheet.Name);
        }

        [Fact]
        public async Task TryGetSheetReturnsFalseWhenNotFound()
        {
            await using var ms = await TypedWorkbook.BuildAsync();
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            Assert.False(reader.TryGetSheet("DoesNotExist", out _));
        }

        [Fact]
        public async Task SheetsIndexerReturnsTheSheetAtThatIndex()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("First", []), ("Second", []));
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            Assert.Equal("Second", reader.Sheets[1].Name);
        }

        [Fact]
        public async Task NegativeSheetIndexThrows()
        {
            await using var ms = await TypedWorkbook.BuildAsync();
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.Sheets[-1]);
            Assert.Throws<ArgumentOutOfRangeException>(() => ((IExcelWorkbook)reader).SheetAt(-1));
        }

        [Fact]
        public async Task OutOfRangeSheetIndexThrows()
        {
            await using var ms = await TypedWorkbook.BuildAsync();
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.Sheets[reader.SheetCount]);
            Assert.Throws<ArgumentOutOfRangeException>(() => ((IExcelWorkbook)reader).SheetAt(reader.SheetCount));
        }

        [Fact]
        public async Task SheetNavigationWorksThroughFormatAgnosticInterface()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("A", [[11]]),
                ("B", [[22]]));
            await using IExcelWorkbook reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);

            var names = new List<string>();
            var firstCells = new List<int>();
            for (int i = 0; i < reader.SheetCount; i++)
            {
                IExcelSheet sheet = reader.SheetAt(i);
                names.Add(sheet.Name);
                await using var e = sheet.GetEnumerator();
                Assert.True(await e.MoveNextAsync());
                Assert.True(e.Current[0].TryParse(null, out int v));
                firstCells.Add(v);
            }

            Assert.Equal(["A", "B"], names);
            Assert.Equal([11, 22], firstCells);
        }

        [Fact]
        [SuppressMessage("Performance", "CA1859:Use concrete types when possible for improved performance",
            Justification = "The test deliberately exercises CSV through the IExcelWorkbook interface, the contract every format shares.")]
        public void CsvIsExposedAsSingleUnnamedSheet()
        {
            using var ms = new MemoryStream("h\nv\n"u8.ToArray());
            using IExcelWorkbook reader = Excel.FromCsv(ms);

            Assert.Equal(1, reader.SheetCount);
            Assert.Equal("", reader.FirstSheet.Name);
            Assert.Equal(ExcelSheetVisibility.Visible, reader.FirstSheet.Visibility);
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.SheetAt(1));
            Assert.True(reader.TryGetSheet("", out _));
            Assert.False(reader.TryGetSheet("Sheet1", out _));
        }

        [Fact]
        public async Task MultipleSheetsEachHaveDistinctData()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("A", [[11]]),
                ("B", [[22]]));
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            await using var e1 = reader.Sheets[0].GetEnumerator();
            Assert.True(await e1.MoveNextAsync());
            Assert.True(e1.Current[0].TryParse(null, out int v1));
            Assert.Equal(11, v1);

            await using var e2 = reader.Sheets[1].GetEnumerator();
            Assert.True(await e2.MoveNextAsync());
            Assert.True(e2.Current[0].TryParse(null, out int v2));
            Assert.Equal(22, v2);
        }
    }
}
