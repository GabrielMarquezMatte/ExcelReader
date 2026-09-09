using System.Data;
using System.Globalization;
using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class ExcelDataReaderTests
    {
        [Fact]
        public async Task HeaderRowFixesNamesAndFieldCount()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Id", "Name"],
                [1, "alice"]);
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);
            using var data = new ExcelDataReader(reader);

            Assert.Equal(2, data.FieldCount);
            Assert.Equal("Id", data.GetName(0));
            Assert.Equal("Name", data.GetName(1));
            Assert.Equal(0, data.GetOrdinal("id"));   // case-insensitive
            Assert.Equal(1, data.GetOrdinal("Name"));
        }

        [Fact]
        public async Task ReadWalksEveryDataRowAfterTheHeader()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Id"],
                [1],
                [2],
                [3]);
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);
            using var data = new ExcelDataReader(reader);

            var seen = new List<int>();
            while (data.Read())
            {
                seen.Add(data.GetInt32(0));
            }
            Assert.Equal([1, 2, 3], seen);
        }

        [Fact]
        public async Task GetValueMapsEachCellTypeToItsClrType()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Text", "Number", "Flag", "When", "Blank"],
                ["hi", 42, true, new DateTime(2024, 1, 1), null]);
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);
            using var data = new ExcelDataReader(reader);

            Assert.True(data.Read());
            Assert.Equal("hi", data.GetValue(0));
            Assert.Equal(42.0, data.GetValue(1));
            Assert.Equal((object)true, data.GetValue(2));
            Assert.Equal(new DateTime(2024, 1, 1), data.GetValue(3));
            Assert.Equal(DBNull.Value, data.GetValue(4));
            Assert.True(data.IsDBNull(4));
        }

        [Fact]
        public async Task NoHeaderSynthesizesColumnNamesFromFirstRowWidth()
        {
            await using var ms = await TypedWorkbook.BuildAsync([1, 2, 3]);
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);
            using var data = new ExcelDataReader(reader, headerRow: 0);

            Assert.Equal(3, data.FieldCount);
            Assert.Equal("Column0", data.GetName(0));
            Assert.True(data.Read());
            Assert.Equal(1.0, data.GetValue(0));
            Assert.False(data.Read()); // that first row was the only row
        }

        [Fact]
        public async Task HeaderRowBeyondSheetSizeIsAnEmptyResultSet()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Id"]);
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);
            using var data = new ExcelDataReader(reader, headerRow: 5);

            Assert.Equal(0, data.FieldCount);
            Assert.False(data.Read());
        }

        [Fact]
        public async Task GetOrdinalForUnknownNameThrows()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Id"], [1]);
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);
            using var data = new ExcelDataReader(reader);

            Assert.Throws<KeyNotFoundException>(() => data.GetOrdinal("DoesNotExist"));
        }

        [Fact]
        public async Task GetValueBeforeReadThrows()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Id"], [1]);
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);
            using var data = new ExcelDataReader(reader);

            Assert.Throws<InvalidOperationException>(() => data.GetValue(0));
        }

        [Fact]
        public async Task NextResultAlwaysReturnsFalse()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("A", [[1]]),
                ("B", [[2]]));
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);
            using var data = new ExcelDataReader(reader);

            Assert.False(data.NextResult());
        }

        // GetBytes reads a text cell's raw UTF-8 bytes directly (no string round trip for a
        // multi-byte character to trip on) and windows them by fieldOffset/length exactly like
        // IDataRecord.GetBytes documents.
        [Fact]
        public async Task GetBytesReturnsTheCellsUtf8BytesWindowedByOffsetAndLength()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Text"], ["héllo"]);
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);
            using var data = new ExcelDataReader(reader);
            Assert.True(data.Read());

            byte[] expected = Encoding.UTF8.GetBytes("héllo");
            Assert.Equal(expected.Length, data.GetBytes(0, 0, null, 0, 0)); // null buffer: just the length

            byte[] buffer = new byte[expected.Length];
            long read = data.GetBytes(0, 0, buffer, 0, buffer.Length);
            Assert.Equal(expected.Length, read);
            Assert.Equal(expected, buffer);

            // Windowed: skip the 2-byte 'é' (UTF-8 bytes 1-2), read the remaining 3 bytes ("llo").
            byte[] window = new byte[3];
            long windowRead = data.GetBytes(0, 3, window, 0, window.Length);
            Assert.Equal(3, windowRead);
            Assert.Equal("llo"u8.ToArray(), window);
        }

        // A numeric cell's bytes never touched Cell.Value (see Cell.GetString's own hasNumber
        // branch) — GetBytes has to format the number itself instead of just slicing Value.
        [Fact]
        public async Task GetBytesFormatsANumericCellInsteadOfReturningNothing()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Number"], [42]);
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);
            using var data = new ExcelDataReader(reader);
            Assert.True(data.Read());

            byte[] buffer = new byte[8];
            long read = data.GetBytes(0, 0, buffer, 0, buffer.Length);
            Assert.Equal("42"u8.ToArray(), buffer.AsSpan(0, (int)read).ToArray());
        }

        [Fact]
        public async Task GetBytesOnADbNullCellReturnsZero()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Blank"], [null]);
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);
            using var data = new ExcelDataReader(reader);
            Assert.True(data.Read());

            Assert.Equal(0, data.GetBytes(0, 0, null, 0, 0));
            byte[] buffer = new byte[4];
            Assert.Equal(0, data.GetBytes(0, 0, buffer, 0, buffer.Length));
        }

        [Fact]
        public async Task GetCharsReturnsTheCellsTextWindowedByOffsetAndLength()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Text"], ["hello world"]);
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);
            using var data = new ExcelDataReader(reader);
            Assert.True(data.Read());

            Assert.Equal(11, data.GetChars(0, 0, null, 0, 0)); // null buffer: just the length

            char[] window = new char[5];
            long read = data.GetChars(0, 6, window, 0, window.Length);
            Assert.Equal(5, read);
            Assert.Equal("world", new string(window));
        }

        [Fact]
        public async Task DataTableLoadPopulatesFromTheSheet()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Id", "Name"],
                [1, "alice"],
                [2, "bob"]);
            await using IExcelRowReader reader = await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken);
            using var data = new ExcelDataReader(reader);

            var table = new DataTable();
            table.Load(data);

            Assert.Equal(2, table.Rows.Count);
            Assert.Equal(["Id", "Name"], table.Columns.Cast<DataColumn>().Select(c => c.ColumnName), StringComparer.Ordinal);
            Assert.Equal(1.0, Convert.ToDouble(table.Rows[0]["Id"], CultureInfo.InvariantCulture));
            Assert.Equal("alice", table.Rows[0]["Name"]);
            Assert.Equal(2.0, Convert.ToDouble(table.Rows[1]["Id"], CultureInfo.InvariantCulture));
            Assert.Equal("bob", table.Rows[1]["Name"]);
        }
    }
}
