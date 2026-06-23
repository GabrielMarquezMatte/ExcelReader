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
        public void StringPropertyIsMapped()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c></row>""" +
                """<row r="2"><c r="A2" t="inlineStr"><is><t>Alice</t></is></c></row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Alice", result[0].Name);
        }

        [Fact]
        public void IntPropertyIsMapped()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Age</t></is></c></row>""" +
                """<row r="2"><c r="A2"><v>42</v></c></row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(42, result[0].Age);
        }

        [Fact]
        public void DoublePropertyIsMapped()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Score</t></is></c></row>""" +
                """<row r="2"><c r="A2"><v>95.5</v></c></row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(95.5, result[0].Score);
        }

        [Fact]
        public void DecimalPropertyIsMapped()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Balance</t></is></c></row>""" +
                """<row r="2"><c r="A2"><v>12345.67</v></c></row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(12345.67m, result[0].Balance);
        }

        [Fact]
        public void BoolPropertyTrueIsMapped()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Active</t></is></c></row>""" +
                """<row r="2"><c r="A2" t="b"><v>1</v></c></row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.True(result[0].Active);
        }

        [Fact]
        public void BoolPropertyFalseIsMapped()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Active</t></is></c></row>""" +
                """<row r="2"><c r="A2" t="b"><v>0</v></c></row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.False(result[0].Active);
        }

        [Fact]
        public void DateTimePropertyIsMapped()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>BirthDate</t></is></c></row>""" +
                """<row r="2"><c r="A2" s="1"><v>25569</v></c></row>""",
                styles: DateStyles);
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(Jan1970, result[0].BirthDate);
        }

        [Fact]
        public void AllBasicTypesAreMappedTogether()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1">""" +
                """<c r="A1" t="inlineStr"><is><t>Name</t></is></c>""" +
                """<c r="B1" t="inlineStr"><is><t>Age</t></is></c>""" +
                """<c r="C1" t="inlineStr"><is><t>Score</t></is></c>""" +
                """<c r="D1" t="inlineStr"><is><t>Active</t></is></c>""" +
                """<c r="E1" t="inlineStr"><is><t>Balance</t></is></c>""" +
                """</row>""" +
                """<row r="2">""" +
                """<c r="A2" t="inlineStr"><is><t>Bob</t></is></c>""" +
                """<c r="B2"><v>35</v></c>""" +
                """<c r="C2"><v>88.25</v></c>""" +
                """<c r="D2" t="b"><v>1</v></c>""" +
                """<c r="E2"><v>500.00</v></c>""" +
                """</row>""");
            using var reader = Excel.From(ms);
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
        public void ExcelColumnAttributeOverridesPropertyName()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1">""" +
                """<c r="A1" t="inlineStr"><is><t>First Name</t></is></c>""" +
                """<c r="B1" t="inlineStr"><is><t>Last Name</t></is></c>""" +
                """<c r="C1" t="inlineStr"><is><t>Count</t></is></c>""" +
                """</row>""" +
                """<row r="2">""" +
                """<c r="A2" t="inlineStr"><is><t>John</t></is></c>""" +
                """<c r="B2" t="inlineStr"><is><t>Doe</t></is></c>""" +
                """<c r="C2"><v>5</v></c>""" +
                """</row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<AttributeRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("John", result[0].FirstName);
            Assert.Equal("Doe", result[0].LastName);
            Assert.Equal(5, result[0].Count);
        }

        [Fact]
        public void HeaderMatchingPropertyNameNotAttributeNameIsIgnored()
        {
            // "FirstName" matches the property name but ExcelColumn says "First Name" — should not match
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>FirstName</t></is></c></row>""" +
                """<row r="2"><c r="A2" t="inlineStr"><is><t>Jane</t></is></c></row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<AttributeRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Null(result[0].FirstName);
        }

        [Fact]
        public void PropertyWithoutAttributeMatchesByName()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Count</t></is></c></row>""" +
                """<row r="2"><c r="A2"><v>7</v></c></row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<AttributeRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(7, result[0].Count);
        }

        // --- Config: ColumnNameComparer ---

        [Fact]
        public void DefaultComparerIsOrdinalIgnoreCase()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>name</t></is></c></row>""" +
                """<row r="2"><c r="A2" t="inlineStr"><is><t>Alice</t></is></c></row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Alice", result[0].Name);
        }

        [Fact]
        public void OrdinalComparerRejectsCaseMismatch()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>name</t></is></c></row>""" +
                """<row r="2"><c r="A2" t="inlineStr"><is><t>Alice</t></is></c></row>""");
            using var reader = Excel.From(ms);
            var config = new ExcelParserConfig { ColumnNameComparer = StringComparer.Ordinal };
            var result = new ExcelParser<PersonRow>(config).Parse(reader).ToList();
            Assert.Single(result);
            Assert.Null(result[0].Name);
        }

        [Fact]
        public void OrdinalComparerAcceptsExactCaseMatch()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c></row>""" +
                """<row r="2"><c r="A2" t="inlineStr"><is><t>Alice</t></is></c></row>""");
            using var reader = Excel.From(ms);
            var config = new ExcelParserConfig { ColumnNameComparer = StringComparer.Ordinal };
            var result = new ExcelParser<PersonRow>(config).Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Alice", result[0].Name);
        }

        // --- Config: HeaderRow ---

        [Fact]
        public void HeaderRowTwoSkipsFirstRow()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Ignored</t></is></c></row>""" +
                """<row r="2"><c r="A2" t="inlineStr"><is><t>Name</t></is></c></row>""" +
                """<row r="3"><c r="A3" t="inlineStr"><is><t>Alice</t></is></c></row>""");
            using var reader = Excel.From(ms);
            var config = new ExcelParserConfig { HeaderRow = 2 };
            var result = new ExcelParser<PersonRow>(config).Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Alice", result[0].Name);
        }

        [Fact]
        public void HeaderRowThreeSkipsTwoRows()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Row1</t></is></c></row>""" +
                """<row r="2"><c r="A2" t="inlineStr"><is><t>Row2</t></is></c></row>""" +
                """<row r="3"><c r="A3" t="inlineStr"><is><t>Name</t></is></c></row>""" +
                """<row r="4"><c r="A4" t="inlineStr"><is><t>Bob</t></is></c></row>""");
            using var reader = Excel.From(ms);
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
        public void ExtraColumnsAreIgnored()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1">""" +
                """<c r="A1" t="inlineStr"><is><t>Name</t></is></c>""" +
                """<c r="B1" t="inlineStr"><is><t>UnknownColumn</t></is></c>""" +
                """</row>""" +
                """<row r="2">""" +
                """<c r="A2" t="inlineStr"><is><t>Carol</t></is></c>""" +
                """<c r="B2" t="inlineStr"><is><t>extra</t></is></c>""" +
                """</row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Carol", result[0].Name);
        }

        [Fact]
        public void MissingColumnsKeepDefault()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c></row>""" +
                """<row r="2"><c r="A2" t="inlineStr"><is><t>Dave</t></is></c></row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Dave", result[0].Name);
            Assert.Equal(0, result[0].Age);
            Assert.Equal(0.0, result[0].Score);
            Assert.False(result[0].Active);
            Assert.Equal(0m, result[0].Balance);
        }

        [Fact]
        public void ColumnOrderIsIrrelevant()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1">""" +
                """<c r="A1" t="inlineStr"><is><t>Age</t></is></c>""" +
                """<c r="B1" t="inlineStr"><is><t>Name</t></is></c>""" +
                """</row>""" +
                """<row r="2">""" +
                """<c r="A2"><v>25</v></c>""" +
                """<c r="B2" t="inlineStr"><is><t>Eve</t></is></c>""" +
                """</row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Eve", result[0].Name);
            Assert.Equal(25, result[0].Age);
        }

        // --- Multiple rows ---

        [Fact]
        public void MultipleRowsHaveNoStateBleed()
        {
            const int count = 12;
            var sb = new System.Text.StringBuilder();
            sb.Append("""<row r="1"><c r="A1" t="inlineStr"><is><t>Age</t></is></c></row>""");
            for (int i = 0; i < count; i++)
            {
                int r = i + 2;
                sb.Append($"""<row r="{r}"><c r="A{r}"><v>{r * 10}</v></c></row>""");
            }
            using var ms = WorkbookBuilder.Build(sb.ToString());
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Equal(count, result.Count);
            for (int i = 0; i < count; i++)
            {
                Assert.Equal((i + 2) * 10, result[i].Age);
            }
        }

        // --- Empty sheet ---

        [Fact]
        public void EmptySheetYieldsNoRows()
        {
            using var ms = WorkbookBuilder.Build("");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Empty(result);
        }

        [Fact]
        public void HeaderOnlyYieldsNoRows()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c></row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Empty(result);
        }

        // --- Nullable types ---

        [Fact]
        public void NullableIntFilledCellYieldsValue()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Quantity</t></is></c></row>""" +
                """<row r="2"><c r="A2"><v>99</v></c></row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<NullableRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(99, result[0].Quantity);
        }

        [Fact]
        public void NullableIntMissingCellYieldsNull()
        {
            // Quantity exists in header at col A; data row only has a cell at col B (unrelated) → col A is a gap
            using var ms = WorkbookBuilder.Build(
                """<row r="1">""" +
                """<c r="A1" t="inlineStr"><is><t>Quantity</t></is></c>""" +
                """<c r="B1" t="inlineStr"><is><t>Rate</t></is></c>""" +
                """</row>""" +
                """<row r="2">""" +
                """<c r="B2"><v>1.5</v></c>""" +
                """</row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<NullableRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Null(result[0].Quantity);
            Assert.Equal(1.5, result[0].Rate);
        }

        [Fact]
        public void NullableDoubleFilledCellYieldsValue()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Rate</t></is></c></row>""" +
                """<row r="2"><c r="A2"><v>3.14</v></c></row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<NullableRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(3.14, result[0].Rate);
        }

        [Fact]
        public void NullableDateTimeFilledCellYieldsValue()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>EventDate</t></is></c></row>""" +
                """<row r="2"><c r="A2" s="1"><v>25569</v></c></row>""",
                styles: DateStyles);
            using var reader = Excel.From(ms);
            var result = new ExcelParser<NullableRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(Jan1970, result[0].EventDate);
        }

        [Fact]
        public void NullableDateTimeMissingCellYieldsNull()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1">""" +
                """<c r="A1" t="inlineStr"><is><t>EventDate</t></is></c>""" +
                """<c r="B1" t="inlineStr"><is><t>Rate</t></is></c>""" +
                """</row>""" +
                """<row r="2">""" +
                """<c r="B2"><v>5.0</v></c>""" +
                """</row>""",
                styles: DateStyles);
            using var reader = Excel.From(ms);
            var result = new ExcelParser<NullableRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Null(result[0].EventDate);
        }

        // --- Struct support ---

        [Fact]
        public void StructRowIsSupported()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1">""" +
                """<c r="A1" t="inlineStr"><is><t>X</t></is></c>""" +
                """<c r="B1" t="inlineStr"><is><t>Y</t></is></c>""" +
                """<c r="C1" t="inlineStr"><is><t>Label</t></is></c>""" +
                """</row>""" +
                """<row r="2">""" +
                """<c r="A2"><v>1.5</v></c>""" +
                """<c r="B2"><v>2.5</v></c>""" +
                """<c r="C2" t="inlineStr"><is><t>Point</t></is></c>""" +
                """</row>""" +
                """<row r="3">""" +
                """<c r="A3"><v>3.0</v></c>""" +
                """<c r="B3"><v>4.0</v></c>""" +
                """<c r="C3" t="inlineStr"><is><t>Vector</t></is></c>""" +
                """</row>""");
            using var reader = Excel.From(ms);
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
        public void TwoStructRowsAreDistinct()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>X</t></is></c></row>""" +
                """<row r="2"><c r="A2"><v>10.0</v></c></row>""" +
                """<row r="3"><c r="A3"><v>20.0</v></c></row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<MeasurementRow>().Parse(reader).ToList();
            Assert.Equal(2, result.Count);
            Assert.Equal(10.0, result[0].X);
            Assert.Equal(20.0, result[1].X);
        }

        // --- Date cells ---

        [Fact]
        public void DateCellWithStandardDateStyleIsConverted()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>BirthDate</t></is></c></row>""" +
                """<row r="2"><c r="A2" s="1"><v>25569</v></c></row>""",
                styles: DateStyles);
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(Jan1970, result[0].BirthDate);
        }

        [Fact]
        public void Date1904IsHandledCorrectly()
        {
            // 1904 date system: serial 0 = Jan 1, 1904
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
        public void NonNumericValueInIntColumnKeepsDefault()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Age</t></is></c></row>""" +
                """<row r="2"><c r="A2" t="inlineStr"><is><t>not-a-number</t></is></c></row>""");
            using var reader = Excel.From(ms);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(0, result[0].Age);
        }

        [Fact]
        public void ParseFailureDoesNotThrowException()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Age</t></is></c></row>""" +
                """<row r="2"><c r="A2" t="inlineStr"><is><t>N/A</t></is></c></row>""");
            using var reader = Excel.From(ms);
            var ex = Record.Exception(() => new ExcelParser<PersonRow>().Parse(reader).ToList());
            Assert.Null(ex);
        }

        // --- Async parity ---

        [Fact]
        public async Task AsyncParseMatchesSyncParse()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1">""" +
                """<c r="A1" t="inlineStr"><is><t>Name</t></is></c>""" +
                """<c r="B1" t="inlineStr"><is><t>Age</t></is></c>""" +
                """</row>""" +
                """<row r="2">""" +
                """<c r="A2" t="inlineStr"><is><t>Alice</t></is></c>""" +
                """<c r="B2"><v>30</v></c>""" +
                """</row>""" +
                """<row r="3">""" +
                """<c r="A3" t="inlineStr"><is><t>Bob</t></is></c>""" +
                """<c r="B3"><v>25</v></c>""" +
                """</row>""");

            List<PersonRow> syncResult;
            List<PersonRow> asyncResult = [];

            using (var reader = Excel.From(ms))
            {
                syncResult = new ExcelParser<PersonRow>().Parse(reader).ToList();
            }

            ms.Position = 0;
            using (var reader = Excel.From(ms))
            {
                await foreach (var row in new ExcelParser<PersonRow>().ParseAsync(reader))
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
            var sb = new System.Text.StringBuilder();
            sb.Append("""<row r="1"><c r="A1" t="inlineStr"><is><t>Age</t></is></c></row>""");
            for (int i = 0; i < count; i++)
            {
                int r = i + 2;
                sb.Append($"""<row r="{r}"><c r="A{r}"><v>{i + 1}</v></c></row>""");
            }
            using var ms = WorkbookBuilder.Build(sb.ToString());
            using var reader = Excel.From(ms);
            var result = new List<PersonRow>();
            await foreach (var row in new ExcelParser<PersonRow>().ParseAsync(reader))
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
