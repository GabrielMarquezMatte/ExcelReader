using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    // Covers [ExcelRequired]: the column must be present in the header row, validated as soon as the
    // header is read. Enforces column presence only, not per-row values.
    public class RequiredTests
    {
        private sealed class Row
        {
            [ExcelRequired]
            public int Id { get; set; }

            [ExcelRequired]
            [ExcelColumn("FullName")]
            public required string Name { get; set; }

            public string? Note { get; set; } // optional
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
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            Row row = new ExcelParser<Row>().Parse(reader).Single();

            Assert.Equal(1, row.Id);
            Assert.Equal("Alice", row.Name);
            Assert.Equal("hi", row.Note);
        }

        [Fact]
        public async Task OptionalColumnMayBeAbsent()
        {
            // Note is not required; its absence must not throw.
            await using var ms = await TypedWorkbook.BuildAsync(["Id", "FullName"], [2, "Bob"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            Row row = new ExcelParser<Row>().Parse(reader).Single();

            Assert.Equal(2, row.Id);
            Assert.Equal("Bob", row.Name);
            Assert.Null(row.Note);
        }

        [Fact]
        public async Task MissingRequiredColumnThrowsListingAll()
        {
            // Neither required header is present (Name's alias is "FullName", not "Name").
            await using var ms = await TypedWorkbook.BuildAsync(["Note"], ["x"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
                () => new ExcelParser<Row>().Parse(reader).ToList());

            Assert.Contains("Id", ex.Message, StringComparison.Ordinal);
            Assert.Contains("FullName", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task RequiredColumnMatchedByAliasSucceeds()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Id", "FullName"], [3, "Carol"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            Row row = new ExcelParser<Row>().Parse(reader).Single();

            Assert.Equal("Carol", row.Name);
        }

        [Fact]
        public async Task EmptyValueInRequiredColumnThrowsWithRowNumber()
        {
            // Header row 1, data rows 2 and 3; row 3's Code cell is blank.
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("S1", [["Code"], ["A1"], [null]]));
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            var enumerator = new ExcelParser<ValueRow>().Parse(reader).GetEnumerator();
            Assert.True(enumerator.MoveNext()); // row 2 ok
            Assert.Equal("A1", enumerator.Current.Code);

            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() => enumerator.MoveNext());
            Assert.Contains("Code", ex.Message, StringComparison.Ordinal);
            Assert.Contains("row 3", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task RequiredColumnWithUnparseableValueThrowsAsIfMissing()
        {
            // "Id" is [ExcelRequired] and its cell is present and non-empty ("oops"), but "oops" isn't
            // a valid int — F3: this must fail the same way an empty required cell would, instead of
            // silently leaving Id at its default (0).
            await using var ms = await TypedWorkbook.BuildAsync(["Id", "FullName"], ["oops", "Alice"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
                () => new ExcelParser<Row>().Parse(reader).ToList());
            Assert.Contains("Id", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task ThrowOnParseFailureReportsColumnAndRawValueEvenForRequiredColumn()
        {
            // ThrowOnParseFailure takes priority over the "missing required value" message: the caller
            // asked for the more specific diagnostic (row/column/raw text), not the generic one.
            await using var ms = await TypedWorkbook.BuildAsync(["Id", "FullName"], ["oops", "Alice"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var config = new ExcelParserConfig { ThrowOnParseFailure = true };

            ExcelParseException ex = Assert.Throws<ExcelParseException>(
                () => new ExcelParser<Row>(config).Parse(reader).ToList());
            Assert.Equal("Id", ex.ColumnName);
            Assert.Equal("oops", ex.RawValue);
        }

        [Fact]
        public async Task AbsentCellInRequiredColumnThrows()
        {
            // Two columns in the header; the second data row omits the Code cell entirely (short row).
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("S1", [["Note", "Code"], ["n1", "c1"], ["n2"]]));
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
                () => new ExcelParser<TwoColRow>().Parse(reader).ToList());
            Assert.Contains("Code", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task AllowEmptyRequiresColumnButPermitsBlankValues()
        {
            await using var ms = await TypedWorkbook.BuildMultiSheetAsync(
                ("S1", [["Code"], ["A1"], [null]]));
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            List<PresenceOnlyRow> rows = [.. new ExcelParser<PresenceOnlyRow>().Parse(reader)];

            Assert.Equal(2, rows.Count);
            Assert.Equal("A1", rows[0].Code);
            Assert.Null(rows[1].Code); // blank allowed
        }

        [Fact]
        public async Task RequiredPropertyWithNoParserThrows()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Thing"], ["v"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
                () => new ExcelParser<UnsupportedRequiredRow>().Parse(reader).ToList());

            Assert.Contains("ExcelRequired", ex.Message, StringComparison.Ordinal);
        }
    }
}
