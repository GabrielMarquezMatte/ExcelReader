using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Tests
{
    public class CoverageEdgeTests
    {
        private sealed class UnsupportedRow
        {
            public Guid Id { get; set; } = Guid.NewGuid();
            public Guid? OptionalId { get; set; } = Guid.NewGuid();
            public string? Name { get; set; }
        }

        private sealed class NullableBoolRow
        {
            public bool? Active { get; set; }
        }

        private sealed class DateRow
        {
            public DateTime Date { get; set; } = DateTime.MinValue;
            public DateTime? OptionalDate { get; set; }
        }

        private sealed class DateOnlyRow
        {
            public DateOnly Date { get; set; }
            public DateOnly? OptionalDate { get; set; }
        }

        [Fact]
        public async Task UnsupportedParserPropertiesAreIgnored()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["Id", "OptionalId", "Name"],
                ["not-a-guid", "also-not-a-guid", "kept"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            UnsupportedRow row = new ExcelParser<UnsupportedRow>().Parse(reader).Single();

            Assert.NotEqual(Guid.Empty, row.Id);
            Assert.NotNull(row.OptionalId);
            Assert.Equal("kept", row.Name);
        }

        [Fact]
        public async Task NullableBoolParserMapsTruthyText()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Active"], ["TRUE"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            NullableBoolRow row = new ExcelParser<NullableBoolRow>().Parse(reader).Single();

            Assert.True(row.Active);
        }

        [Fact]
        public async Task InvalidDateValuesKeepDefaults()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Date", "OptionalDate"], ["nope", "still-nope"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            DateRow row = new ExcelParser<DateRow>().Parse(reader).Single();

            Assert.Equal(DateTime.MinValue, row.Date);
            Assert.Null(row.OptionalDate);
        }

        [Fact]
        public async Task DateOnlyParserMapsDateCellsAndKeepsDefaultsForEmptyOrInvalidCells()
        {
            var date = new DateTime(2024, 6, 27, 0, 0, 0, DateTimeKind.Unspecified);
            await using var filled = await TypedWorkbook.BuildAsync(["Date", "OptionalDate"], [date, date]);
            await using var filledReader = await Excel.FromAsync(filled, ct: TestContext.Current.CancellationToken);

            DateOnlyRow parsed = new ExcelParser<DateOnlyRow>().Parse(filledReader).Single();
            Assert.Equal(new DateOnly(2024, 6, 27), parsed.Date);
            Assert.Equal(new DateOnly(2024, 6, 27), parsed.OptionalDate);

            await using var empty = await TypedWorkbook.BuildAsync(["Date", "OptionalDate"], [new Gap(), new Gap()]);
            await using var emptyReader = await Excel.FromAsync(empty, ct: TestContext.Current.CancellationToken);
            DateOnlyRow emptyParsed = new ExcelParser<DateOnlyRow>().Parse(emptyReader).Single();
            Assert.Equal(default, emptyParsed.Date);
            Assert.Null(emptyParsed.OptionalDate);

            await using var invalid = await TypedWorkbook.BuildAsync(["Date", "OptionalDate"], ["nope", "still-nope"]);
            await using var invalidReader = await Excel.FromAsync(invalid, ct: TestContext.Current.CancellationToken);
            DateOnlyRow invalidParsed = new ExcelParser<DateOnlyRow>().Parse(invalidReader).Single();
            Assert.Equal(default, invalidParsed.Date);
            Assert.Null(invalidParsed.OptionalDate);
        }

        [Fact]
        public async Task InterfacesReturnFreshEnumerators()
        {
            await using var ms = await TypedWorkbook.BuildAsync(["Name"], ["Alice"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var enumerable = new ExcelParser<UnsupportedRow>().Parse(reader);

            using IEnumerator<UnsupportedRow> typed = enumerable.GetEnumerator();
            System.Collections.IEnumerator untyped = ((System.Collections.IEnumerable)enumerable).GetEnumerator();
            using var disposable = Assert.IsAssignableFrom<IDisposable>(untyped);

            Assert.True(typed.MoveNext());
            Assert.True(untyped.MoveNext());
        }

        [Fact]
        public async Task AsyncParserWithHeaderAfterSkippedRowsAndNoDataReturnsNoRows()
        {
            await using var ms = await TypedWorkbook.BuildAsync(
                ["ignored"],
                ["Name"]);
            await using var reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            var parser = new ExcelParser<UnsupportedRow>(new ExcelParserConfig { HeaderRow = 2 });

            var rows = new List<UnsupportedRow>();
            await foreach (UnsupportedRow row in parser.ParseAsync(reader, TestContext.Current.CancellationToken))
            {
                rows.Add(row);
            }

            Assert.Empty(rows);
        }

        [Fact]
        public async Task AsyncReaderHandlesSelfClosingAndUnknownElements()
        {
            await using MemoryStream ms = WorkbookBuilder.Build(
                """
                <row r="1"><c r="A1" t="inlineStr"><is><t>value</t></is></c><ext/></row>
                <row r="2"><c r="A2"/><odd><x/></odd><c r="B2"><v>42</v></c></row>
                """);
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            await using XlsxReader.Enumerator e = await reader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);

            Assert.True(await e.MoveNextAsync());
            Assert.Equal("value", e.Current[0].GetString());
            Assert.True(await e.MoveNextAsync());
            Assert.Equal(CellType.Empty, e.Current[0].Type);
            Assert.Equal(42, int.Parse(e.Current[1].GetString(), CultureInfo.InvariantCulture));
            Assert.False(await e.MoveNextAsync());
        }

        [Fact]
        public async Task AsyncReaderLargeCellForcesSlowRefillPaths()
        {
            string text = new('x', 80_000);
            await using MemoryStream ms = WorkbookBuilder.Build(
                $"""<row r="1"><c r="A1" t="inlineStr"><is><t>{text}</t></is></c></row>""");
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            await using XlsxReader.Enumerator e = await reader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);

            Assert.True(await e.MoveNextAsync());
            Assert.Equal(80_000, e.Current[0].Value.Length);
            Assert.False(await e.MoveNextAsync());
        }

        [Fact]
        public async Task AsyncReaderEmptyAndMalformedSheetsReturnFalse()
        {
            await using MemoryStream empty = WorkbookBuilder.Build("");
            await using XlsxReader emptyReader = await Excel.FromAsync(empty, ct: TestContext.Current.CancellationToken);
            await using XlsxReader.Enumerator emptyRows = await emptyReader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);

            Assert.False(await emptyRows.MoveNextAsync());

            await using MemoryStream malformed = BuildRawWorkbook(
                """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>""",
                """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>""",
                ("xl/worksheets/sheet1.xml", "<worksheet><sheetData><dimension"));
            await using XlsxReader malformedReader = await Excel.FromAsync(malformed, ct: TestContext.Current.CancellationToken);
            await using XlsxReader.Enumerator malformedRows = await malformedReader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);

            Assert.False(await malformedRows.MoveNextAsync());

            await using MemoryStream noTags = BuildRawWorkbook(
                """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>""",
                """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>""",
                ("xl/worksheets/sheet1.xml", "plain text"));
            await using XlsxReader noTagsReader = await Excel.FromAsync(noTags, ct: TestContext.Current.CancellationToken);
            await using XlsxReader.Enumerator noTagsRows = await noTagsReader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);

            Assert.False(await noTagsRows.MoveNextAsync());
        }

        [Fact]
        public async Task AsyncReaderHandlesHugeRowOpenTagAndMissingRowClose()
        {
            string padding = new('x', 70_000);
            await using MemoryStream hugeOpenTag = WorkbookBuilder.Build(
                $"""<row r="1" custom="{padding}"><c r="A1"><v>7</v></c></row>""");
            await using XlsxReader reader = await Excel.FromAsync(hugeOpenTag, ct: TestContext.Current.CancellationToken);
            await using XlsxReader.Enumerator rows = await reader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);

            Assert.True(await rows.MoveNextAsync());
            Assert.Equal("7", rows.Current[0].GetString());
            Assert.False(await rows.MoveNextAsync());

            await using MemoryStream missingClose = BuildRawWorkbook(
                """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>""",
                """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>""",
                ("xl/worksheets/sheet1.xml", """<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c>"""));
            await using XlsxReader malformedReader = await Excel.FromAsync(missingClose, ct: TestContext.Current.CancellationToken);
            await using XlsxReader.Enumerator malformedRows = await malformedReader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);

            Assert.True(await malformedRows.MoveNextAsync());
            Assert.False(await malformedRows.MoveNextAsync());
        }

        [Fact]
        public async Task AsyncReaderMissingRowOpenTagReturnsEmptyRow()
        {
            await using MemoryStream ms = BuildRawWorkbook(
                """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>""",
                """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>""",
                ("xl/worksheets/sheet1.xml", "<worksheet><sheetData><row r=\"1\""));
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            await using XlsxReader.Enumerator rows = await reader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);

            Assert.True(await rows.MoveNextAsync());
            Assert.Equal(0, rows.Current.ColumnCount);
            Assert.False(await rows.MoveNextAsync());
        }

        [Fact]
        public void SyncReaderMissingRowOpenTagReturnsEmptyRow()
        {
            using MemoryStream ms = BuildRawWorkbook(
                """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>""",
                """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>""",
                ("xl/worksheets/sheet1.xml", "<worksheet><sheetData><row r=\"1\""));
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator rows = reader.GetEnumerator();

            Assert.True(rows.MoveNext());
            Assert.Equal(0, rows.Current.ColumnCount);
            Assert.False(rows.MoveNext());
        }

        [Fact]
        public void ReaderSkipsUnknownTopLevelElementsAndStopsOnEnd()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """
                <dimension ref="A1"/>
                <row r="1"><c r="A1"><v>1</v></c><ignored><child/></ignored></row>
                """);
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(1, int.Parse(e.Current[0].GetString(), CultureInfo.InvariantCulture));
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void ReaderSkipsCDataBetweenRowsAndHandlesMalformedCellValues()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """
                <row r="1"><c r="A1" t="q"><v></v></c><c><v>5</v></c></row>
                <![CDATA[ignored > marker]]>
                <row r="2"><c r="A2" t="inlineStr"><is><t>after</t></is></c></row>
                """);
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator rows = reader.GetEnumerator();

            Assert.True(rows.MoveNext());
            Assert.Equal(string.Empty, rows.Current[0].GetString());
            Assert.Equal("5", rows.Current[1].GetString());
            Assert.True(rows.MoveNext());
            Assert.Equal("after", rows.Current[0].GetString());
            Assert.False(rows.MoveNext());
        }

        [Fact]
        public void ReaderMalformedRowsAndCellsExitGracefully()
        {
            using MemoryStream missingCellClose = WorkbookBuilder.Build(
                """<row r="1"><c r="A1"><v>1</v></row>""");
            using XlsxReader reader1 = Excel.From(missingCellClose);
            using XlsxReader.Enumerator e1 = reader1.GetEnumerator();

            Assert.True(e1.MoveNext());
            Assert.Equal(CellType.Empty, e1.Current[0].Type);
            Assert.False(e1.MoveNext());

            using MemoryStream missingRowClose = BuildRawWorkbook(
                """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>""",
                """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>""",
                ("xl/worksheets/sheet1.xml", """<worksheet><sheetData><row r="1"><ignored"""));
            using XlsxReader reader2 = Excel.From(missingRowClose);
            using XlsxReader.Enumerator e2 = reader2.GetEnumerator();

            Assert.True(e2.MoveNext());
            Assert.False(e2.MoveNext());
        }

        [Fact]
        public void SharedStringsWithLowUniqueCountGrowOffsetsArray()
        {
            using MemoryStream ms = BuildWorkbookWithRawSharedStrings(
                """<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="1"><si><t>A</t></si><si><t>B</t></si><si><t>C</t></si></sst>""",
                """<row r="1"><c r="A1" t="s"><v>2</v></c></row>""");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("C", e.Current[0].GetString());
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void SharedStringsHandleMalformedAndSelfClosingItems()
        {
            using MemoryStream ms = BuildWorkbookWithRawSharedStrings(
                """<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="4"><si/><si><t>A</t></si><si><t>unterminated</si></sst>""",
                """<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>""");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(string.Empty, e.Current[0].GetString());
            Assert.Equal("A", e.Current[1].GetString());
            Assert.Equal(string.Empty, e.Current[2].GetString());
        }

        [Fact]
        public void SharedStringsCoverMalformedTextRunsAndNumericEntities()
        {
            using MemoryStream ms = BuildWorkbookWithRawSharedStrings(
                """<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t/></si><si><r/></si><si><t>bad &#xZZ;</t></si><si><t</si></sst>""",
                """<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c></row>""");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(string.Empty, e.Current[0].GetString());
            Assert.Equal(string.Empty, e.Current[1].GetString());
            Assert.Equal("bad &#xZZ;", e.Current[2].GetString());
            Assert.Equal(string.Empty, e.Current[3].GetString());
        }

        [Fact]
        public void MalformedWorkbookPartsThrowOrReturnEmptyData()
        {
            using MemoryStream noSheets = BuildRawWorkbook(
                """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>""",
                """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>""",
                []);

            Assert.Throws<InvalidDataException>(() => Excel.From(noSheets));
        }

        [Fact]
        public void RelationshipTargetsAreNormalized()
        {
            using MemoryStream ms = BuildRawWorkbook(
                """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Absolute" sheetId="1" r:id="rId1"/><sheet name="AlreadyRooted" sheetId="2" r:id="rId2"/></sheets></workbook>""",
                """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="/xl/worksheets/sheet1.xml"/><Relationship Id="rId2" Type="x" Target="xl/worksheets/sheet2.xml"/></Relationships>""",
                ("xl/worksheets/sheet1.xml", """<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"""),
                ("xl/worksheets/sheet2.xml", """<worksheet><sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData></worksheet>"""));
            using XlsxReader reader = Excel.From(ms);

            Assert.Equal(2, reader.SheetCount);
            Assert.Equal("Absolute", reader.SheetName);
            reader.MoveToSheet(1);
            Assert.Equal("AlreadyRooted", reader.SheetName);
        }

        [Fact]
        public async Task AsyncSharedStringsLoadOnlyOnce()
        {
            await using MemoryStream ms = BuildWorkbookWithRawSharedStrings(
                """<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>A</t></si></sst>""",
                """<row r="1"><c r="A1" t="s"><v>0</v></c></row>""");
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);

            await using (XlsxReader.Enumerator first = await reader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken))
            {
                Assert.True(await first.MoveNextAsync());
            }
            await using XlsxReader.Enumerator second = await reader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);
            Assert.True(await second.MoveNextAsync());
        }

        [Fact]
        public async Task AsyncOpenDisposesStreamOnInvalidWorkbookWhenNotLeavingOpen()
        {
            var ms = new TrackingMemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                Write(zip, "xl/workbook.xml", """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>""");
                Write(zip, "xl/_rels/workbook.xml.rels", """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>""");
            }
            ms.Position = 0;

            await Assert.ThrowsAsync<InvalidDataException>(async () =>
                await Excel.FromAsync(ms, leaveOpen: false, ct: TestContext.Current.CancellationToken));
            Assert.True(ms.Disposed);
        }

        [Fact]
        public void StylesMalformedAndEscapedFormatsFallBackSafely()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" s="1"><v>123</v></c><c r="B1" s="2"><v>456</v></c></row>""",
                styles:
                """
                <styleSheet>
                  <numFmts count="3">
                    <numFmt numFmtId="200" formatCode="&quot;year&quot;"/>
                    <numFmt numFmtId="201" formatCode="[red]0"/>
                    <numFmt numFmtId="202" formatCode="\y"/>
                  </numFmts>
                  <cellStyleXfs><xf numFmtId="14"/></cellStyleXfs>
                  <cellXfs><xf numFmtId="0"/><xf numFmtId="200"/><xf numFmtId="201"/></cellXfs>
                </styleSheet>
                """);
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Row row = e.Current;

            Assert.Equal(CellType.Number, row[0].Type);
            Assert.Equal(CellType.Number, row[1].Type);
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void StylesMalformedRegionsAndUnclosedLiteralsAreIgnored()
        {
            using MemoryStream missingCellXfs = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" s="1"><v>1</v></c></row>""",
                styles: """<styleSheet><numFmts><numFmt numFmtId="200" formatCode="mm-dd-yy"/></numFmts></styleSheet>""");
            using XlsxReader reader1 = Excel.From(missingCellXfs);
            using XlsxReader.Enumerator e1 = reader1.GetEnumerator();
            Assert.True(e1.MoveNext());
            Assert.Equal(CellType.Number, e1.Current[0].Type);

            using MemoryStream malformedCellXfs = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" s="0"><v>1</v></c><c r="B1" s="1"><v>2</v></c></row>""",
                styles:
                """
                <styleSheet>
                  <numFmts>
                    <numFmt numFmtId="200" formatCode="&quot;unterminated"/>
                    <numFmt numFmtId="201" formatCode="[unterminated"/>
                  </numFmts>
                  <cellXfs><xf numFmtId="200"/><xf numFmtId="201"/>
                </styleSheet>
                """);
            using XlsxReader reader2 = Excel.From(malformedCellXfs);
            using XlsxReader.Enumerator e2 = reader2.GetEnumerator();
            Assert.True(e2.MoveNext());
            Assert.Equal(CellType.Number, e2.Current[0].Type);
            Assert.Equal(CellType.Number, e2.Current[1].Type);
        }

        [Fact]
        public async Task WriterCoversNullableBoolAndLargeColumnNames()
        {
            await using var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken))
            {
                await wb.StartAsync(TestContext.Current.CancellationToken);
                XlsxSheetWriter sheet = wb.AddSheet("Wide");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write((bool?)null);
                    row.Write((bool?)true);
                    row.Skip(700);
                    row.Write(123);
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }

            ms.Position = 0;
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());

            Assert.Equal(CellType.Empty, e.Current[0].Type);
            Assert.Equal("1", e.Current[1].GetString());
            Assert.Equal("123", e.Current[702].GetString());
            Assert.False(e.MoveNext());
        }

        [Fact]
        public async Task WriterStateErrorsCoverStartedAndEndedBranches()
        {
            await using XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(new MemoryStream(), ct: TestContext.Current.CancellationToken);

            await wb.StartAsync(TestContext.Current.CancellationToken);
            await Assert.ThrowsAsync<InvalidOperationException>(async () =>
                await wb.StartAsync(TestContext.Current.CancellationToken));
            XlsxSheetWriter sheet = wb.AddSheet("S1");
            Assert.Throws<InvalidOperationException>(() => wb.AddSheet("S2"));
            await sheet.StartAsync(TestContext.Current.CancellationToken);
            await sheet.EndAsync(TestContext.Current.CancellationToken);
            await wb.EndAsync(TestContext.Current.CancellationToken);

            Assert.Throws<ObjectDisposedException>(() => wb.AddSheet("S2"));
        }

        [Fact]
        public async Task SheetWriterStateErrorsCoverStartedAndEndedBranches()
        {
            await using XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(new MemoryStream(), ct: TestContext.Current.CancellationToken);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            XlsxSheetWriter sheet = wb.AddSheet("S1");

            await Assert.ThrowsAsync<InvalidOperationException>(async () =>
                await sheet.EndAsync(TestContext.Current.CancellationToken));
            await sheet.StartAsync(TestContext.Current.CancellationToken);
            await Assert.ThrowsAsync<InvalidOperationException>(async () =>
                await sheet.StartAsync(TestContext.Current.CancellationToken));
            await sheet.EndAsync(TestContext.Current.CancellationToken);
            await Assert.ThrowsAsync<ObjectDisposedException>(async () =>
                await sheet.StartRowAsync(TestContext.Current.CancellationToken));
        }

        [Fact]
        public void CellToStringAndApostropheEscapingAreCovered()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>it&apos;s</t></is></c></row>""");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Cell cell = e.Current[0];

            Assert.Equal("it's", cell.ToString());
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void XlsxXmlDecodeHandlesLiteralMarkupAndUnterminatedCData()
        {
            Span<byte> dest = stackalloc byte[64];

            int written = XlsxXml.Decode("a <tag &unknown;"u8, dest);
            Assert.Equal("a <tag &unknown;", Encoding.UTF8.GetString(dest[..written]));

            written = XlsxXml.Decode("<![CDATA[unterminated"u8, dest);
            Assert.Equal("unterminated", Encoding.UTF8.GetString(dest[..written]));
        }

        [Fact]
        public void ColumnNameWritesUtf8Form()
        {
            Span<byte> bytes = stackalloc byte[3];

            int written = ColumnName.Write(bytes, 0);
            Assert.Equal(1, written);
            Assert.Equal("A", Encoding.ASCII.GetString(bytes[..written]));

            written = ColumnName.Write(bytes, 25);
            Assert.Equal(1, written);
            Assert.Equal("Z", Encoding.ASCII.GetString(bytes[..written]));

            written = ColumnName.Write(bytes, 26);
            Assert.Equal(2, written);
            Assert.Equal("AA", Encoding.ASCII.GetString(bytes[..written]));

            written = ColumnName.Write(bytes, 701);
            Assert.Equal(2, written);
            Assert.Equal("ZZ", Encoding.ASCII.GetString(bytes[..written]));

            written = ColumnName.Write(bytes, 702);
            Assert.Equal(3, written);
            Assert.Equal("AAA", Encoding.ASCII.GetString(bytes[..written]));

            written = ColumnName.Write(bytes, 16_383);
            Assert.Equal(3, written);
            Assert.Equal("XFD", Encoding.ASCII.GetString(bytes[..written]));
        }

        [Fact]
        public void LimitChecksThrowNamedExceptions()
        {
            var options = new ExcelReaderOptions
            {
                MaxCellBytes = 4,
                MaxSharedStringBytes = 8,
            };

            ExcelLimitExceededException shared = Assert.Throws<ExcelLimitExceededException>(() =>
                LimitChecks.ThrowIfOverSharedStringLimit(options, 9));
            Assert.Equal(nameof(ExcelReaderOptions.MaxSharedStringBytes), shared.LimitName);

            ExcelLimitExceededException array = Assert.Throws<ExcelLimitExceededException>(() =>
                LimitChecks.NextBufferSize(0, nameof(ExcelReaderOptions.MaxCellBytes), Array.MaxLength, Array.MaxLength));
            Assert.Equal("ArrayMaxLength", array.LimitName);
        }

        [Fact]
        public void MissingWorkbookBytesThrowNoSheets()
        {
            using MemoryStream ms = BuildRawWorkbook(
                workbookXml: null,
                relsXml: null,
                worksheets: []);

            Assert.Throws<InvalidDataException>(() => Excel.From(ms));
        }

        private static MemoryStream BuildWorkbookWithRawSharedStrings(string sharedXml, string sheetRows)
        {
            return BuildRawWorkbook(
                """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>""",
                """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>""",
                [("xl/worksheets/sheet1.xml", $"""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{sheetRows}</sheetData></worksheet>"""),
                 ("xl/sharedStrings.xml", sharedXml)]);
        }

        private static MemoryStream BuildRawWorkbook(string? workbookXml, string? relsXml, params (string Path, string Xml)[] worksheets)
        {
            var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                if (workbookXml is not null)
                {
                    Write(zip, "xl/workbook.xml", workbookXml);
                }
                if (relsXml is not null)
                {
                    Write(zip, "xl/_rels/workbook.xml.rels", relsXml);
                }
                foreach ((string path, string xml) in worksheets)
                {
                    Write(zip, path, xml);
                }
            }
            ms.Position = 0;
            return ms;
        }

        private static void Write(ZipArchive zip, string name, string content)
        {
            using Stream s = zip.CreateEntry(name).Open();
            byte[] bytes = Encoding.UTF8.GetBytes(content);
            s.Write(bytes);
        }

        private sealed class TrackingMemoryStream : MemoryStream
        {
            internal bool Disposed { get; private set; }

            protected override void Dispose(bool disposing)
            {
                Disposed = true;
                base.Dispose(disposing);
            }
        }
    }
}
