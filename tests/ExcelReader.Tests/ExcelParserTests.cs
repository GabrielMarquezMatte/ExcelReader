using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class ExcelParserTests
    {
        private sealed class PersonRow
        {
            public string? Name { get; set; }
            public int Age { get; set; }
            public double Score { get; set; }
            public DateTime BirthDate { get; set; }
            public bool Active { get; set; }
            public decimal Balance { get; set; }
        }

        private sealed class AttributeRow
        {
            [ExcelColumn("First Name")]
            public string? FirstName { get; set; }
            [ExcelColumn("Last Name")]
            public string? LastName { get; set; }
            public int Count { get; set; }
        }

        private struct MeasurementRow
        {
            public double X { get; set; }
            public double Y { get; set; }
            public string? Label { get; set; }
        }

        private sealed class NullableRow
        {
            public int? Quantity { get; set; }
            public DateTime? EventDate { get; set; }
            public double? Rate { get; set; }
        }

        private const string DateStyles =
            """<styleSheet><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>""";

        // OADate 25569 = January 1, 1970
        private static readonly DateTime Jan1970 = DateTime.FromOADate(25569);

        // --- Basic property mapping ---

        [Fact]
        public async Task StringPropertyIsMapped()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name"], ["Alice"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Alice", result[0].Name);
        }

        [Fact]
        public async Task IntPropertyIsMapped()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Age"], [42]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(42, result[0].Age);
        }

        [Fact]
        public async Task DoublePropertyIsMapped()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Score"], [95.5]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(95.5, result[0].Score);
        }

        [Fact]
        public async Task DecimalPropertyIsMapped()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Balance"], [12345.67m]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(12345.67m, result[0].Balance);
        }

        [Fact]
        public async Task BoolPropertyTrueIsMapped()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Active"], [true]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.True(result[0].Active);
        }

        [Fact]
        public async Task BoolPropertyFalseIsMapped()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Active"], [false]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.False(result[0].Active);
        }

        [Fact]
        public async Task AllBasicTypesAreMappedTogether()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Name", "Age", "Score", "Active", "Balance"],
                ["Bob", 35, 88.25, true, 500.00m]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Bob", result[0].Name);
            Assert.Equal(35, result[0].Age);
            Assert.Equal(88.25, result[0].Score);
            Assert.True(result[0].Active);
            Assert.Equal(500.00m, result[0].Balance);
        }

        // --- ExcelColumnAttribute ---

        [Fact]
        public async Task ExcelColumnAttributeOverridesPropertyName()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["First Name", "Last Name", "Count"],
                ["John", "Doe", 5]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<AttributeRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("John", result[0].FirstName);
            Assert.Equal("Doe", result[0].LastName);
            Assert.Equal(5, result[0].Count);
        }

        [Fact]
        public async Task HeaderMatchingPropertyNameNotAttributeNameIsIgnored()
        {
            // "FirstName" matches the property name but ExcelColumn says "First Name" — should not match
            await using var ms = await TypedWorkbook.BuildAsync(["FirstName"], ["Jane"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<AttributeRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Null(result[0].FirstName);
        }

        [Fact]
        public async Task PropertyWithoutAttributeMatchesByName()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Count"], [7]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<AttributeRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(7, result[0].Count);
        }

        // --- Config: ColumnNameComparer ---

        [Fact]
        public async Task DefaultComparerIsOrdinalIgnoreCase()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["name"], ["Alice"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Alice", result[0].Name);
        }

        [Fact]
        public async Task OrdinalComparerRejectsCaseMismatch()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["name"], ["Alice"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var config = new ExcelParserConfig { ColumnNameComparer = StringComparer.Ordinal };
            var result = new ExcelParser<PersonRow>(config).Parse(reader).ToList();
            Assert.Single(result);
            Assert.Null(result[0].Name);
        }

        [Fact]
        public async Task OrdinalComparerAcceptsExactCaseMatch()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name"], ["Alice"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var config = new ExcelParserConfig { ColumnNameComparer = StringComparer.Ordinal };
            var result = new ExcelParser<PersonRow>(config).Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Alice", result[0].Name);
        }

        // --- Config: HeaderRow ---

        [Fact]
        public async Task HeaderRowTwoSkipsFirstRow()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Ignored"],
                ["Name"],
                ["Alice"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var config = new ExcelParserConfig { HeaderRow = 2 };
            var result = new ExcelParser<PersonRow>(config).Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Alice", result[0].Name);
        }

        [Fact]
        public async Task HeaderRowThreeSkipsTwoRows()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Row1"],
                ["Row2"],
                ["Name"],
                ["Bob"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var config = new ExcelParserConfig { HeaderRow = 3 };
            var result = new ExcelParser<PersonRow>(config).Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Bob", result[0].Name);
        }

        // --- Config: invalid config ---

        [Fact]
        public void ZeroHeaderRowThrows()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                new ExcelParser<PersonRow>(new ExcelParserConfig { HeaderRow = 0 }));
        }

        [Fact]
        public void NegativeHeaderRowThrows()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                new ExcelParser<PersonRow>(new ExcelParserConfig { HeaderRow = -1 }));
        }

        // --- Robustness ---

        [Fact]
        public async Task ExtraColumnsAreIgnored()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Name", "UnknownColumn"],
                ["Carol", "extra"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Carol", result[0].Name);
        }

        [Fact]
        public async Task MissingColumnsKeepDefault()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name"], ["Dave"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Dave", result[0].Name);
            Assert.Equal(0, result[0].Age);
            Assert.Equal(0.0, result[0].Score);
            Assert.False(result[0].Active);
            Assert.Equal(0m, result[0].Balance);
        }

        [Fact]
        public async Task ColumnOrderIsIrrelevant()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Age", "Name"],
                [25, "Eve"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Eve", result[0].Name);
            Assert.Equal(25, result[0].Age);
        }

        // --- Multiple rows ---

        [Fact]
        public async Task MultipleRowsHaveNoStateBleed()
        {
            const int count = 12;
            var rows = new object?[count + 1][];
            rows[0] = ["Age"];
            for (int i = 0; i < count; i++)
            {
                rows[i + 1] = [(i + 2) * 10];
            }
            await using var ms = await TypedWorkbook.BuildAsync(rows);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Equal(count, result.Count);
            for (int i = 0; i < count; i++)
            {
                Assert.Equal((i + 2) * 10, result[i].Age);
            }
        }

        // --- Empty sheet ---

        [Fact]
        public async Task EmptySheetYieldsNoRows()
        {
            await using var ms = await TypedWorkbook.BuildAsync();
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Empty(result);
        }

        [Fact]
        public async Task HeaderOnlyYieldsNoRows()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Empty(result);
        }

        // --- Nullable types ---

        [Fact]
        public async Task NullableIntFilledCellYieldsValue()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Quantity"], [99]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<NullableRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(99, result[0].Quantity);
        }

        [Fact]
        public async Task NullableIntMissingCellYieldsNull()
        {
            // Quantity exists in header at col A; data row leaves col A as a gap.
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Quantity", "Rate"],
                [new Gap(), 1.5]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<NullableRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Null(result[0].Quantity);
            Assert.Equal(1.5, result[0].Rate);
        }

        [Fact]
        public async Task NullableDoubleFilledCellYieldsValue()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Rate"], [3.14]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<NullableRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(3.14, result[0].Rate);
        }

        [Fact]
        public async Task NullableDateTimeFilledCellYieldsValue()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["EventDate"], [Jan1970]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<NullableRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(Jan1970, result[0].EventDate);
        }

        [Fact]
        public async Task NullableDateTimeMissingCellYieldsNull()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["EventDate", "Rate"],
                [new Gap(), 5.0]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<NullableRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Null(result[0].EventDate);
        }

        // --- Struct support ---

        [Fact]
        public async Task StructRowIsSupported()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["X", "Y", "Label"],
                [1.5, 2.5, "Point"],
                [3.0, 4.0, "Vector"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<MeasurementRow>().Parse(reader).ToList();
            Assert.Equal(2, result.Count);
            Assert.Equal(1.5, result[0].X);
            Assert.Equal(2.5, result[0].Y);
            Assert.Equal("Point", result[0].Label);
            Assert.Equal(3.0, result[1].X);
            Assert.Equal(4.0, result[1].Y);
            Assert.Equal("Vector", result[1].Label);
        }

        [Fact]
        public async Task TwoStructRowsAreDistinct()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["X"],
                [10.0],
                [20.0]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<MeasurementRow>().Parse(reader).ToList();
            Assert.Equal(2, result.Count);
            Assert.Equal(10.0, result[0].X);
            Assert.Equal(20.0, result[1].X);
        }

        // --- Date cells ---

        [Fact]
        public async Task DateCellWithStandardDateStyleIsConverted()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["BirthDate"], [Jan1970]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(Jan1970, result[0].BirthDate);
        }

        [Fact]
        public void Date1904IsHandledCorrectly()
        {
            // 1904 date system: serial 0 = Jan 1, 1904. WorkbookWriter only emits the
            // 1900 system, so this fixture stays on the raw-XML builder.
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>BirthDate</t></is></c></row>""" +
                """<row r="2"><c r="A2" s="1"><v>0</v></c></row>""",
                styles: DateStyles,
                date1904: true);
            using var reader = Excel.From(ms);
            Assert.True(reader.IsDate1904);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(new DateTime(1904, 1, 1, 0, 0, 0, DateTimeKind.Unspecified), result[0].BirthDate);
        }

        // --- Parse failure (no exceptions) ---

        [Fact]
        public async Task NonNumericValueInIntColumnKeepsDefault()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Age"], ["not-a-number"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(0, result[0].Age);
        }

        [Fact]
        public async Task ParseFailureDoesNotThrowException()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Age"], ["N/A"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var ex = Record.Exception(() => new ExcelParser<PersonRow>().Parse(reader).ToList());
            Assert.Null(ex);
        }

        // --- Async parity ---

        [Fact]
        public async Task AsyncParseMatchesSyncParse()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Name", "Age"],
                ["Alice", 30],
                ["Bob", 25]);

            List<PersonRow> syncResult;
            List<PersonRow> asyncResult = [];

            await using (var reader = Excel.From(ms))
            {
                syncResult = new ExcelParser<PersonRow>().Parse(reader).ToList();
            }

            ms.Position = 0;
            await using (var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken))
            {
                await foreach (var row in new ExcelParser<PersonRow>().ParseAsync(reader, TestContext.Current.CancellationToken))
                {
                    asyncResult.Add(row);
                }
            }

            Assert.Equal(syncResult.Count, asyncResult.Count);
            for (int i = 0; i < syncResult.Count; i++)
            {
                Assert.Equal(syncResult[i].Name, asyncResult[i].Name);
                Assert.Equal(syncResult[i].Age, asyncResult[i].Age);
            }
        }

        [Fact]
        public async Task AsyncParseYieldsAllRows()
        {
            const int count = 5;
            var rows = new object?[count + 1][];
            rows[0] = ["Age"];
            for (int i = 0; i < count; i++)
            {
                rows[i + 1] = [i + 1];
            }
            await using var ms = await TypedWorkbook.BuildAsync(rows);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var result = new List<PersonRow>();
            await foreach (var row in new ExcelParser<PersonRow>().ParseAsync(reader, TestContext.Current.CancellationToken))
            {
                result.Add(row);
            }
            Assert.Equal(count, result.Count);
            for (int i = 0; i < count; i++)
            {
                Assert.Equal(i + 1, result[i].Age);
            }
        }

        // --- Shared strings in header ---

        [Fact]
        public void SharedStringHeaderIsResolved()
        {
            // Shared-strings table is a raw-XML feature WorkbookWriter does not emit.
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v>0</v></c></row>""" +
                """<row r="2"><c r="A2" t="inlineStr"><is><t>Alice</t></is></c></row>""",
                sharedStrings: "<si><t>Name</t></si>");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Alice", result[0].Name);
        }
    }
}
