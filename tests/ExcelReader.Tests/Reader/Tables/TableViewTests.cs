using ExcelReader.Core;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Tests.Reader.Tables
{
    public class TableViewTests
    {
        private sealed class FakeRows(int[] rowIndexes) : IExcelRowEnumerator, IRowIndexedEnumerator
        {
            private int _next;

            public int MoveNextCalls { get; private set; }

            public bool Disposed { get; private set; }

            public bool Enabled { get; private set; }

            public int RowIndex { get; private set; } = -1;

            public Row Current
            {
                get
                {
                    return default;
                }
            }

            public void EnableRowIndex()
            {
                Enabled = true;
            }

            public bool MoveNext()
            {
                MoveNextCalls++;
                if (_next == rowIndexes.Length)
                {
                    return false;
                }
                RowIndex = rowIndexes[_next++];
                return true;
            }

            public ValueTask<bool> MoveNextAsync()
            {
                return new ValueTask<bool>(MoveNext());
            }

            public void Dispose()
            {
                Disposed = true;
            }

            public ValueTask DisposeAsync()
            {
                Disposed = true;
                return ValueTask.CompletedTask;
            }
        }

        [Fact]
        public void BoundedRowsStopAtTheFirstRowPastTheTable()
        {
            FakeRows rows = new([0, 3, 4, 5, 6, 7, 8, 9]);
            using TableRowEnumerator table = new(rows, firstRow: 3, lastDataRow: 6, firstColumn: 0, lastColumn: 0);

            int yielded = 0;
            while (table.MoveNext())
            {
                yielded++;
            }

            Assert.True(rows.Enabled);
            Assert.Equal(4, yielded);
            Assert.Equal(6, rows.MoveNextCalls);
            Assert.False(table.MoveNext());
            Assert.Equal(6, rows.MoveNextCalls);
        }

        [Fact]
        public async Task BoundedRowsStopAtTheFirstRowPastTheTableAsync()
        {
            FakeRows rows = new([0, 3, 4, 5, 6, 7, 8, 9]);
            await using TableRowEnumerator table = new(rows, firstRow: 3, lastDataRow: 6, firstColumn: 0, lastColumn: 0);

            int yielded = 0;
            while (await table.MoveNextAsync())
            {
                yielded++;
            }

            Assert.Equal(4, yielded);
            Assert.Equal(6, rows.MoveNextCalls);
        }

        public static TheoryData<int[]> MalformedRowIndexes { get; } = new()
        {
            new[] { 0, -1 },
            new[] { 0, ExcelLimits.MaxRows },
            new[] { 0, 2, 2 },
            new[] { 0, 3, 2 },
        };

        [Theory]
        [MemberData(nameof(MalformedRowIndexes))]
        public void MalformedRowIndexThrowsInvalidData(int[] rowIndexes)
        {
            using TableRowEnumerator table = new(new FakeRows(rowIndexes), firstRow: 0, lastDataRow: 10, firstColumn: 0, lastColumn: 0);

            Assert.Throws<InvalidDataException>(() =>
            {
                while (table.MoveNext())
                {
                }
            });
        }

        [Theory]
        [MemberData(nameof(MalformedRowIndexes))]
        public async Task MalformedRowIndexThrowsInvalidDataAsync(int[] rowIndexes)
        {
            await using TableRowEnumerator table = new(new FakeRows(rowIndexes), firstRow: 0, lastDataRow: 10, firstColumn: 0, lastColumn: 0);

            await Assert.ThrowsAsync<InvalidDataException>(async () =>
            {
                while (await table.MoveNextAsync())
                {
                }
            });
        }

        [Fact]
        public void DisposingTheViewDisposesTheSheetEnumerator()
        {
            FakeRows rows = new([0]);
            TableRowEnumerator table = new(rows, firstRow: 0, lastDataRow: 0, firstColumn: 0, lastColumn: 0);

            table.Dispose();

            Assert.True(rows.Disposed);
        }

        [Theory]
        [MemberData(nameof(TableMetadataTests.Fixtures), MemberType = typeof(TableMetadataTests))]
        public void SalesViewYieldsHeaderAndDataRowsCutToTheTable(string file)
        {
            using IExcelWorkbook workbook = Excel.Open(TableMetadataTests.FixturePath(file));

            Assert.Equal(
                ["2,3,4:Product,Qty", "2,3,4:Pen,10", "2,3,4:Notebook,4", "2,3,4:Pencil,25"],
                Render(Table(workbook, "Sales").AsSheet(), 2, 3));
        }

        [Theory]
        [MemberData(nameof(TableMetadataTests.Fixtures), MemberType = typeof(TableMetadataTests))]
        public void StockViewIgnoresTheNeighbouringTableAndCells(string file)
        {
            using IExcelWorkbook workbook = Excel.Open(TableMetadataTests.FixturePath(file));

            Assert.Equal(
                ["6,7:Product,Qty", "6,7:Pen,100", "6,7:Notebook,40", "6,7:Pencil,250"],
                Render(Table(workbook, "Stock").AsSheet(), 6, 7));
        }

        [Theory]
        [MemberData(nameof(TableMetadataTests.Fixtures), MemberType = typeof(TableMetadataTests))]
        public void HeaderlessViewYieldsOnlyDataRows(string file)
        {
            using IExcelWorkbook workbook = Excel.Open(TableMetadataTests.FixturePath(file));

            Assert.Equal(["0,1:a,1", "0,1:b,2", "0,1:c,3"], Render(Table(workbook, "NoHeader").AsSheet(), 0, 1));
        }

        [Theory]
        [MemberData(nameof(TableMetadataTests.Fixtures), MemberType = typeof(TableMetadataTests))]
        public void ViewReportsItsOwningSheet(string file)
        {
            using IExcelWorkbook workbook = Excel.Open(TableMetadataTests.FixturePath(file));
            IExcelSheet view = Table(workbook, "NoStyle").AsSheet();

            Assert.Equal(("Extra", 1, ExcelSheetVisibility.Visible), (view.Name, view.Index, view.Visibility));
            Assert.Equal(["3,4:Name,Value", "3,4:x,1", "3,4:y,2"], Render(view, 3, 4));
        }

        [Theory]
        [MemberData(nameof(TableMetadataTests.Fixtures), MemberType = typeof(TableMetadataTests))]
        public async Task AsyncViewOverAStreamYieldsTheSameRows(string file)
        {
            await using IExcelWorkbook workbook = await Excel.OpenAsync(
                File.OpenRead(TableMetadataTests.FixturePath(file)), leaveOpen: false, ct: TestContext.Current.CancellationToken);
            IExcelSheet view = Table(workbook, "Sales").AsSheet();

            List<string> rows = [];
            await using IExcelRowEnumerator enumerator = view.GetAsyncEnumerator(TestContext.Current.CancellationToken);
            while (await enumerator.MoveNextAsync())
            {
                rows.Add(RenderRow(enumerator.Current, 2, 3));
            }

            Assert.Equal(["2,3,4:Product,Qty", "2,3,4:Pen,10", "2,3,4:Notebook,4", "2,3,4:Pencil,25"], rows);
        }

        private static ExcelTable Table(IExcelWorkbook workbook, string name)
        {
            Assert.True(workbook.TryGetTable(name, out ExcelTable? table));
            return table;
        }

        private static List<string> Render(IExcelSheet sheet, int firstShown, int secondShown)
        {
            List<string> rows = [];
            using IExcelRowEnumerator enumerator = sheet.GetEnumerator();
            while (enumerator.MoveNext())
            {
                rows.Add(RenderRow(enumerator.Current, firstShown, secondShown));
            }
            return rows;
        }

        private static string RenderRow(Row row, int firstShown, int secondShown)
        {
            List<int> columns = [];
            foreach (RowCell cell in row.Cells)
            {
                columns.Add(cell.ColumnIndex);
            }
            return $"{string.Join(',', columns)}:{row[firstShown].GetString()},{row[secondShown].GetString()}";
        }
    }
}
