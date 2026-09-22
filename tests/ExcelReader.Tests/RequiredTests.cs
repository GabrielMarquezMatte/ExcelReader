using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class RequiredTests
    {
        private sealed class Row
        {
            [ExcelRequired]
            public int Id { get; set; }

            [ExcelRequired]
            [ExcelColumn("FullName")]
            public required string Name { get; set; }

            public string? Note { get; set; }
        }

        private sealed class ValueRow
        {
            [ExcelRequired]
            public string? Code { get; set; }
        }

        private sealed class PresenceOnlyRow
        {
            [ExcelRequired(AllowEmpty = true)]
            public string? Code { get; set; }
        }

        private sealed class TwoColRow
        {
            public string? Note { get; set; }

            [ExcelRequired]
            public required string Code { get; set; }
        }

        private sealed class CustomValue
        {
            public string? Raw { get; set; }
        }

        private sealed class UnsupportedRequiredRow
        {
            [ExcelRequired]
            public CustomValue Thing { get; set; } = new();
        }

        [Fact]
        public async Task RequiredColumnsPresentParseNormally()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Id", "FullName", "Note"],
                [1, "Alice", "hi"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            Row row = new ExcelParser<Row>().Parse(reader).Single();

            Assert.Equal(1, row.Id);
            Assert.Equal("Alice", row.Name);
            Assert.Equal("hi", row.Note);
        }

        [Fact]
        public async Task OptionalColumnMayBeAbsent()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Id", "FullName"], [2, "Bob"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            Row row = new ExcelParser<Row>().Parse(reader).Single();

            Assert.Equal(2, row.Id);
            Assert.Equal("Bob", row.Name);
            Assert.Null(row.Note);
        }

        [Fact]
        public async Task MissingRequiredColumnThrowsListingAll()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Note"], ["x"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            ExcelParseException ex = Assert.Throws<ExcelParseException>(
                () => new ExcelParser<Row>().Parse(reader).ToList());

            Assert.Contains("Id", ex.Message, StringComparison.Ordinal);
            Assert.Contains("FullName", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task RequiredColumnMatchedByAliasSucceeds()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Id", "FullName"], [3, "Carol"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            Row row = new ExcelParser<Row>().Parse(reader).Single();

            Assert.Equal("Carol", row.Name);
        }

        [Fact]
        public async Task EmptyValueInRequiredColumnThrowsWithRowNumber()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("S1", [["Code"], ["A1"], [null]]));
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            var enumerator = new ExcelParser<ValueRow>().Parse(reader).GetEnumerator();
            Assert.True(enumerator.MoveNext());
            Assert.Equal("A1", enumerator.Current.Code);

            // Advancing a row never parses it, so a bad row can be skipped; the failure surfaces on Current.
            Assert.True(enumerator.MoveNext());
            ExcelParseException ex = Assert.Throws<ExcelParseException>(() => enumerator.Current);
            Assert.Contains("Code", ex.Message, StringComparison.Ordinal);
            Assert.Contains("row 3", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task RequiredColumnWithUnparseableValueThrowsAsIfMissing()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Id", "FullName"], ["oops", "Alice"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            ExcelParseException ex = Assert.Throws<ExcelParseException>(
                () => new ExcelParser<Row>().Parse(reader).ToList());
            Assert.Contains("Id", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task ThrowOnParseFailureReportsColumnAndRawValueEvenForRequiredColumn()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Id", "FullName"], ["oops", "Alice"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            var config = new ExcelParserConfig { ThrowOnParseFailure = true };

            ExcelParseException ex = Assert.Throws<ExcelParseException>(
                () => new ExcelParser<Row>(config).Parse(reader).ToList());
            Assert.Equal("Id", ex.ColumnName);
            Assert.Equal("oops", ex.RawValue);
        }

        [Fact]
        public async Task AbsentCellInRequiredColumnThrows()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("S1", [["Note", "Code"], ["n1", "c1"], ["n2"]]));
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            ExcelParseException ex = Assert.Throws<ExcelParseException>(
                () => new ExcelParser<TwoColRow>().Parse(reader).ToList());
            Assert.Contains("Code", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task AllowEmptyRequiresColumnButPermitsBlankValues()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("S1", [["Code"], ["A1"], [null]]));
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            List<PresenceOnlyRow> rows = [.. new ExcelParser<PresenceOnlyRow>().Parse(reader)];

            Assert.Equal(2, rows.Count);
            Assert.Equal("A1", rows[0].Code);
            Assert.Null(rows[1].Code);
        }

        [Fact]
        public async Task RequiredPropertyWithNoParserThrows()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Thing"], ["v"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);

            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
                () => new ExcelParser<UnsupportedRequiredRow>().Parse(reader).ToList());

            Assert.Contains("ExcelRequired", ex.Message, StringComparison.Ordinal);
        }
    }
}
