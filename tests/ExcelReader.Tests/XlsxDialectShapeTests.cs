using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    // Hand-authored XML fragments mimicking known quirks of specific producers (LibreOffice, openpyxl,
    // Apache POI, etc.) — not files actually exported by them. See RealWorldXlsxCorpusTests for tests
    // against genuine producer-exported binaries.
    public class XlsxDialectShapeTests
    {
        public static IEnumerable<object[]> Fixtures
        {
            get
            {
                yield return
                [
                    new CorpusFixture(
                        "LibreOffice-style whitespace, spans, and single quotes",
                        """
                        <row r="1" spans="1:4">
                            <c r='A1' t='inlineStr'><is><t xml:space="preserve"> leading trailing </t></is></c>
                            <!-- metadata can contain > characters -->
                            <c r='C1'><v>12.5</v></c>
                            <c r='D1' t='b'><v>1</v></c>
                        </row>
                        """,
                        null,
                        [
                            new ExpectedCell(0, 0, " leading trailing ", CellType.ExcelString),
                            new ExpectedCell(0, 2, "12.5", CellType.Number),
                            new ExpectedCell(0, 3, "1", CellType.Boolean),
                        ])
                ];

                yield return
                [
                    new CorpusFixture(
                        "openpyxl-style sparse inline and shared strings",
                        """
                        <row r="1">
                            <c r="B1" t="inlineStr"><is><t>inline &amp; escaped</t></is></c>
                            <c r="E1" t="s"><v>0</v></c>
                        </row>
                        <row r="2">
                            <c r="A2"><v>-3</v></c>
                            <c r="D2" t="inlineStr"><is><t/></is></c>
                        </row>
                        """,
                        "<si><r><t>rich</t></r><r><t> text</t></r></si>",
                        [
                            new ExpectedCell(0, 1, "inline & escaped", CellType.ExcelString),
                            new ExpectedCell(0, 4, "rich text", CellType.ExcelString),
                            new ExpectedCell(1, 0, "-3", CellType.Number),
                            new ExpectedCell(1, 3, "", CellType.ExcelString),
                        ])
                ];

                yield return
                [
                    new CorpusFixture(
                        "Apache POI-style formulas and explicit cell types",
                        """
                        <row r="1">
                            <c r="A1" t="str"><f>CONCAT(&quot;a&quot;,&quot;b&quot;)</f><v>ab</v></c>
                            <c r="B1" t="e"><v>#DIV/0!</v></c>
                            <c r="C1" t="b"><v>0</v></c>
                        </row>
                        """,
                        null,
                        [
                            new ExpectedCell(0, 0, "ab", CellType.Formula),
                            new ExpectedCell(0, 1, "#DIV/0!", CellType.Error),
                            new ExpectedCell(0, 2, "0", CellType.Boolean),
                        ])
                ];

                yield return
                [
                    new CorpusFixture(
                        "Excelize-style single-quoted attributes and CDATA text",
                        """
                        <row r='1'>
                            <c r='A1' t='s'><v>0</v></c>
                            <c r='B1' t='inlineStr'><is><t><![CDATA[raw &amp; <tag>]]></t></is></c>
                        </row>
                        """,
                        "<si><t><![CDATA[shared &amp; <tag>]]></t></si>",
                        [
                            new ExpectedCell(0, 0, "shared &amp; <tag>", CellType.ExcelString),
                            new ExpectedCell(0, 1, "raw &amp; <tag>", CellType.ExcelString),
                        ])
                ];

                yield return
                [
                    new CorpusFixture(
                        "Google Sheets-style extra row markup between cells",
                        """
                        <row r="1" customFormat="0" ht="21">
                            <c r="A1" t="s"><v>0</v></c>
                            <extLst><ext uri="{fixture}"><ignored value="1"/></ext></extLst>
                        </row>
                        <!-- exported comments may sit between rows -->
                        <row r="2">
                            <c r="A2" t="inlineStr"><is><t>next</t></is></c>
                        </row>
                        """,
                        "<si><t>sheet title</t></si>",
                        [
                            new ExpectedCell(0, 0, "sheet title", CellType.ExcelString),
                            new ExpectedCell(1, 0, "next", CellType.ExcelString),
                        ])
                ];
            }
        }

        [Theory]
        [MemberData(nameof(Fixtures))]
        public void SyncReaderHandlesCorpusFixture(CorpusFixture fixture)
        {
            using MemoryStream ms = WorkbookBuilder.Build(fixture.Rows, fixture.SharedStrings);
            using XlsxReader reader = Excel.From(ms);

            Assert.NotEmpty(fixture.Expected);
            AssertExpected(reader.GetEnumerator(), fixture.Expected);
        }

        [Theory]
        [MemberData(nameof(Fixtures))]
        public async Task AsyncReaderHandlesCorpusFixture(CorpusFixture fixture)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            await using MemoryStream ms = WorkbookBuilder.Build(fixture.Rows, fixture.SharedStrings);
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: ct);
            await using XlsxReader.Enumerator rows = await reader.GetAsyncEnumeratorAsync(ct);

            Assert.NotEmpty(fixture.Expected);
            await AssertExpectedAsync(rows, fixture.Expected);
        }

        private static void AssertExpected(XlsxReader.Enumerator rows, ExpectedCell[] expected)
        {
            int next = 0;
            int rowIndex = 0;
            while (rows.MoveNext())
            {
                next = AssertRow(rows.Current, expected, next, rowIndex);
                rowIndex++;
            }

            Assert.Equal(expected.Length, next);
        }

        private static async Task AssertExpectedAsync(
            XlsxReader.Enumerator rows,
            ExpectedCell[] expected)
        {
            int next = 0;
            int rowIndex = 0;
            while (await rows.MoveNextAsync())
            {
                next = AssertRow(rows.Current, expected, next, rowIndex);
                rowIndex++;
            }

            Assert.Equal(expected.Length, next);
        }

        private static int AssertRow(Row row, ExpectedCell[] expected, int next, int rowIndex)
        {
            while (next < expected.Length && expected[next].Row == rowIndex)
            {
                ExpectedCell cell = expected[next++];
                Assert.True(row.ColumnCount > cell.Column);
                Assert.Equal(cell.Type, row[cell.Column].Type);
                Assert.Equal(cell.Value, row[cell.Column].GetString());
            }

            return next;
        }

        public sealed record CorpusFixture(
            string Name,
            string Rows,
            string? SharedStrings,
            ExpectedCell[] Expected)
        {
            public override string ToString()
            {
                return Name;
            }
        }

        public readonly record struct ExpectedCell(int Row, int Column, string Value, CellType Type);
    }
}
