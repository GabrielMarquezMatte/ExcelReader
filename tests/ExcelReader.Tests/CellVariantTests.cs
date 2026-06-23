using System.Globalization;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class CellVariantTests
    {
        [Fact]
        public void BooleanCellHasBooleanType()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>0</v></c></row>""");
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            var row = e.Current;
            Assert.Equal(CellType.Boolean, row[0].Type);
            Assert.Equal("1", row[0].GetString());
            Assert.Equal(CellType.Boolean, row[1].Type);
            Assert.Equal("0", row[1].GetString());
        }

        [Fact]
        public void ErrorCellHasErrorType()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c></row>""");
            using var reader = Excel.From(ms);
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
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Formula, e.Current[0].Type);
            Assert.Equal("Hello", e.Current[0].GetString());
        }

        [Fact]
        public void EmptyWorksheetYieldsNoRows()
        {
            using var ms = WorkbookBuilder.Build("");
            using var reader = Excel.From(ms);
            int count = 0;
            foreach (var _ in reader)
            {
                count++;
            }
            Assert.Equal(0, count);
        }

        [Fact]
        public void SelfClosingRowYieldsZeroColumnCount()
        {
            using var ms = WorkbookBuilder.Build("""<row r="1"/>""");
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(0, e.Current.ColumnCount);
        }

        [Fact]
        public void DecodesQuotAndAposEntities()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>""",
                sharedStrings: """<si><t>say &quot;hi&quot;</t></si><si><t>it&apos;s</t></si>""");
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            var row = e.Current;
            Assert.Equal("say \"hi\"", row[0].GetString());
            Assert.Equal("it's", row[1].GetString());
        }

        [Fact]
        public void DecodesHexNumericEntity()
        {
            // &#x41; = 'A', &#X7A; = 'z' (uppercase X prefix is also accepted)
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v>0</v></c></row>""",
                sharedStrings: """<si><t>&#x41;&#X7A;</t></si>""");
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("Az", e.Current[0].GetString());
        }

        [Fact]
        public void RichTextRunsConcatenated()
        {
            // Two <r> runs per shared string; WriteTextRuns must concatenate both <t> values.
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v>0</v></c></row>""",
                sharedStrings: """<si><r><rPr><b/></rPr><t>Hello</t></r><r><rPr><i/></rPr><t> World</t></r></si>""");
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("Hello World", e.Current[0].GetString());
        }

        [Fact]
        public void DateFormatLettersInsideQuotesIsNotDate()
        {
            // &quot;dd/mm/yyyy&quot; decodes to "dd/mm/yyyy" — every letter is quoted → not a date.
            const string styles =
                """<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;dd/mm/yyyy&quot;"/></numFmts>""" +
                """<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164"/></cellXfs></styleSheet>""";
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" s="1"><v>45292</v></c></row>""", styles: styles);
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Number, e.Current[0].Type);
        }

        [Fact]
        public void DateFormatLettersInsideBracketsIsNotDate()
        {
            // [y] — the letter is inside a bracket section → not a date.
            const string styles =
                """<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="[y]"/></numFmts>""" +
                """<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164"/></cellXfs></styleSheet>""";
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" s="1"><v>45292</v></c></row>""", styles: styles);
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Number, e.Current[0].Type);
        }

        [Fact]
        public void DateFormatBackslashEscapedLetterIsNotDate()
        {
            // \y — the backslash escapes the letter, so it doesn't count → not a date.
            const string styles =
                """<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="\y"/></numFmts>""" +
                """<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164"/></cellXfs></styleSheet>""";
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" s="1"><v>45292</v></c></row>""", styles: styles);
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Number, e.Current[0].Type);
        }

        [Fact]
        public void TryGetDateTimeOnInlineStringReturnsFalse()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>not-a-date</t></is></c></row>""");
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.False(e.Current[0].TryGetDateTime(out _));
        }

        [Fact]
        public void TryGetDateTimeOutOfRangeSerialReturnsFalse()
        {
            // 3000000 exceeds the OADate upper bound of 2958466.
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1"><v>3000000</v></c></row>""");
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.False(e.Current[0].TryGetDateTime(out _));
        }

        [Fact]
        public void LeaveOpenFalseClosesStreamOnDispose()
        {
            using var ms = WorkbookBuilder.Build("");
            using (var reader = Excel.From(ms, leaveOpen: false))
            {
                Assert.Equal(1, reader.SheetCount);
            }
            Assert.False(ms.CanRead);
        }

        [Fact]
        public void LeaveOpenTrueKeepsStreamOpenAfterDispose()
        {
            using var ms = WorkbookBuilder.Build("");
            using (var reader = Excel.From(ms, leaveOpen: true))
            {
                Assert.Equal(1, reader.SheetCount);
            }
            Assert.True(ms.CanRead);
        }

        [Fact]
        public void LargeColumnReferenceAAIsColumn26()
        {
            // AA in base-26: 1*26 + 1 = 27 → zero-based index 26.
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="AA1"><v>99</v></c></row>""");
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            var row = e.Current;
            Assert.Equal(CellType.Empty, row[0].Type);
            Assert.Equal(CellType.Number, row[26].Type);
            Assert.True(row[26].TryParse(null, out int v));
            Assert.Equal(99, v);
        }

        [Fact]
        public void TryParseDoubleSucceeds()
        {
            using var ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1"><v>3.14</v></c></row>""");
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.True(e.Current[0].TryParse(CultureInfo.InvariantCulture, out double d));
            Assert.Equal(3.14, d);
        }

        [Fact]
        public void TryParseLongSucceeds()
        {
            using var ms = WorkbookBuilder.Build(
                $"""<row r="1"><c r="A1"><v>{long.MaxValue}</v></c></row>""");
            using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.True(e.Current[0].TryParse(null, out long l));
            Assert.Equal(long.MaxValue, l);
        }
    }
}
