using System.Globalization;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class CellVariantTests
    {
        [Fact]
        public async Task BooleanCellHasBooleanType()
        {
            await using var ms = await TypedWorkbook.BuildAsync([true, false]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            await using var e = reader.GetEnumerator();
            Assert.True(await e.MoveNextAsync());
            Assert.Equal(CellType.Boolean, e.Current[0].Type);
            Assert.Equal("1", e.Current[0].GetString());
            Assert.Equal(CellType.Boolean, e.Current[1].Type);
            Assert.Equal("0", e.Current[1].GetString());
        }

        [Fact]
        public void ErrorCellHasErrorType()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c></row>""");
            using var reader = Excel.FromXlsx(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Error, e.Current[0].Type);
            Assert.Equal("#DIV/0!", e.Current[0].GetString());
        }

        [Fact]
        public void FormulaCellHasFormulaType()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="str"><v>Hello</v></c></row>""");
            using var reader = Excel.FromXlsx(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Formula, e.Current[0].Type);
            Assert.Equal("Hello", e.Current[0].GetString());
        }

        [Fact]
        public void TryGetDoubleRejectsCommaDecimalInsteadOfMisreadingIt()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="str"><v>1,5</v></c><c r="B1" t="str"><v>1.5</v></c></row>""");
            using var reader = Excel.FromXlsx(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.False(e.Current[0].TryGetDouble(out double ambiguous));
            Assert.Equal(0.0, ambiguous);
            Assert.True(e.Current[1].TryGetDouble(out double unambiguous));
            Assert.Equal(1.5, unambiguous);
        }

        [Fact]
        public async Task EmptyWorksheetYieldsNoRows()
        {
            await using var ms = await TypedWorkbook.BuildAsync();
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            await using var enumerator = await reader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);
            int count = 0;
            while (await enumerator.MoveNextAsync())
            {
                count++;
            }
            Assert.Equal(0, count);
        }

        [Fact]
        public void SelfClosingRowYieldsZeroColumnCount()
        {
            using var ms = WorkbookBuilder.Build("""<row r="1"/>""");
            using var reader = Excel.FromXlsx(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(0, e.Current.ColumnCount);
        }

        [Fact]
        public void DecodesQuotAndAposEntities()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>""",
                sharedStrings: "<si><t>say &quot;hi&quot;</t></si><si><t>it&apos;s</t></si>");
            using var reader = Excel.FromXlsx(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            var row = e.Current;
            Assert.Equal("say \"hi\"", row[0].GetString());
            Assert.Equal("it's", row[1].GetString());
        }

        [Fact]
        public void DecodesHexNumericEntity()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v>0</v></c></row>""",
                sharedStrings: "<si><t>&#x41;&#X7A;</t></si>");
            using var reader = Excel.FromXlsx(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("Az", e.Current[0].GetString());
        }

        [Fact]
        public void RichTextRunsConcatenated()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v>0</v></c></row>""",
                sharedStrings: "<si><r><rPr><b/></rPr><t>Hello</t></r><r><rPr><i/></rPr><t> World</t></r></si>");
            using var reader = Excel.FromXlsx(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("Hello World", e.Current[0].GetString());
        }

        [Fact]
        public void DateFormatLettersInsideQuotesIsNotDate()
        {
            const string styles =
                """<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;dd/mm/yyyy&quot;"/></numFmts>""" +
                """<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164"/></cellXfs></styleSheet>""";
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" s="1"><v>45292</v></c></row>""", styles: styles);
            using var reader = Excel.FromXlsx(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Number, e.Current[0].Type);
        }

        [Fact]
        public void DateFormatLettersInsideBracketsIsNotDate()
        {
            const string styles =
                """<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="[y]"/></numFmts>""" +
                """<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164"/></cellXfs></styleSheet>""";
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" s="1"><v>45292</v></c></row>""", styles: styles);
            using var reader = Excel.FromXlsx(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Number, e.Current[0].Type);
        }

        [Fact]
        public void DateFormatBackslashEscapedLetterIsNotDate()
        {
            const string styles =
                """<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="\y"/></numFmts>""" +
                """<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164"/></cellXfs></styleSheet>""";
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" s="1"><v>45292</v></c></row>""", styles: styles);
            using var reader = Excel.FromXlsx(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Number, e.Current[0].Type);
        }

        [Fact]
        public async Task TryGetDateTimeOnInlineStringReturnsFalse()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["not-a-date"]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            await using var e = reader.GetEnumerator();
            Assert.True(await e.MoveNextAsync());
            Assert.False(e.Current[0].TryGetDateTime(out _));
        }

        [Fact]
        public async Task TryGetDateTimeOutOfRangeSerialReturnsFalse()
        {
            await using var ms = await TypedWorkbook.BuildAsync([3000000]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            await using var e = reader.GetEnumerator();
            Assert.True(await e.MoveNextAsync());
            Assert.False(e.Current[0].TryGetDateTime(out _));
        }

        [Fact]
        public async Task LeaveOpenFalseClosesStreamOnDispose()
        {
            await using var ms = await TypedWorkbook.BuildAsync();
            await using (var reader = await Excel.FromXlsxAsync(ms, leaveOpen: false, ct: TestContext.Current.CancellationToken))
            {
                Assert.Equal(1, reader.SheetCount);
            }
            Assert.False(ms.CanRead);
        }

        [Fact]
        public async Task LeaveOpenTrueKeepsStreamOpenAfterDispose()
        {
            await using var ms = await TypedWorkbook.BuildAsync();
            await using (var reader = await Excel.FromXlsxAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken))
            {
                Assert.Equal(1, reader.SheetCount);
            }
            Assert.True(ms.CanRead);
        }

        [Fact]
        public async Task LargeColumnReferenceAAIsColumn26()
        {
            await using var ms = await TypedWorkbook.BuildAsync([new Gap(26), 99]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            await using var e = reader.GetEnumerator();
            Assert.True(await e.MoveNextAsync());
            Assert.Equal(CellType.Empty, e.Current[0].Type);
            Assert.Equal(CellType.Number, e.Current[26].Type);
            Assert.True(e.Current[26].TryParse(null, out int v));
            Assert.Equal(99, v);
        }

        [Fact]
        public async Task TryParseDoubleSucceeds()
        {
            await using var ms = await TypedWorkbook.BuildAsync([3.14]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            await using var e = reader.GetEnumerator();
            Assert.True(await e.MoveNextAsync());
            Assert.True(e.Current[0].TryParse(CultureInfo.InvariantCulture, out double d));
            Assert.Equal(3.14, d);
        }

        [Fact]
        public async Task TryParseLongSucceeds()
        {
            await using var ms = await TypedWorkbook.BuildAsync([long.MaxValue]);
            await using var reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            await using var e = reader.GetEnumerator();
            Assert.True(await e.MoveNextAsync());
            Assert.True(e.Current[0].TryParse(null, out long l));
            Assert.Equal(long.MaxValue, l);
        }

        [Fact]
        public void Date1904WorkbookReadsCorrectDates()
        {
            const string styles =
                """<styleSheet><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>""";
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" s="1"><v>0</v></c><c r="B1" s="1"><v>1</v></c></row>""",
                styles: styles,
                date1904: true);
            using var reader = Excel.FromXlsx(ms);
            Assert.True(reader.IsDate1904);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            var row = e.Current;
            Assert.Equal(CellType.Date, row[0].Type);
            Assert.True(row[0].TryGetDateTime(reader.IsDate1904, out var d0));
            Assert.Equal(new DateTime(1904, 1, 1, 0, 0, 0, DateTimeKind.Unspecified), d0);
            Assert.True(row[1].TryGetDateTime(reader.IsDate1904, out var d1));
            Assert.Equal(new DateTime(1904, 1, 2, 0, 0, 0, DateTimeKind.Unspecified), d1);
        }

        [Fact]
        public async Task IsDate1904FalseForStandard1900Workbook()
        {
            await using var ms = await TypedWorkbook.BuildAsync();
            await using var reader = Excel.FromXlsx(ms);
            Assert.False(reader.IsDate1904);
        }

        [Fact]
        public void Excel1900SerialsOneToSixtyMapToExcelCalendar()
        {
            const string styles =
                """<styleSheet><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>""";
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="1"><v>59</v></c><c r="C1" s="1"><v>60</v></c></row>""",
                styles: styles);
            using var reader = Excel.FromXlsx(ms);
            using var e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.True(e.Current[0].TryGetDateTime(out DateTime first));
            Assert.True(e.Current[1].TryGetDateTime(out DateTime leapBoundary));
            Assert.True(e.Current[2].TryGetDateTime(out DateTime phantomLeapDay));
            Assert.Equal(new DateTime(1900, 1, 1), first);
            Assert.Equal(new DateTime(1900, 2, 28), leapBoundary);
            Assert.Equal(new DateTime(1900, 2, 28), phantomLeapDay);
        }

        [Fact]
        public void BinaryNumberOutsideDecimalRangeDoesNotThrow()
        {
            using var ms = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1E+30</v></c></row>""");
            using var reader = Excel.FromXlsx(ms);
            using var e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.False(e.Current[0].TryParse(null, out decimal _));
        }
    }
}
