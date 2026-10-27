using System.Globalization;
using System.Text;
using ExcelReader.Core;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Tests.Reader.Xlsx
{
    public class XlsxCanonicalCellTests
    {
        // An element prefix sends every cell through the general path, so the default-namespace sheet,
        // whose cells in Excel's shape take the fast path, must read back the same.
        [Fact]
        public void CellsInExcelsShapeReadTheSameAsThroughTheGeneralPath()
        {
            const int Rows = 2000;
            using var plain = WorkbookBuilder.Build(SheetRows("", Rows), SharedStrings(""), $"<styleSheet>{StylesInner("")}</styleSheet>");
            using var prefixed = WorkbookBuilder.BuildPrefixed("x", SheetRows("x:", Rows), SharedStrings("x:"), StylesInner("x:"));

            List<string> expected = Snapshot(prefixed);
            List<string> actual = Snapshot(plain);

            Assert.Equal(Rows, expected.Count);
            Assert.Equal(expected, actual);
        }

        private static readonly string[] FormulaCells =
        [
            """<c r="A{r}"><f>A{r}+1</f><v>3</v></c>""",
            """<c r="B{r}"><f t="shared" ref="B1:B9" si="0">A1*2</f><v>4.5</v></c>""",
            """<c r="C{r}"><f t="shared" si="0"/><v>6</v></c>""",
            """<c r="D{r}" s="1"><f t="array" ref="D1:D2">{1;2}</f><v>45000</v></c>""",
            """<c r="E{r}" t="str"><f>IF(A1&gt;2,"a&lt;b",'q"')</f><v>a&lt;b</v></c>""",
            """<c r="F{r}" t="str"><f>A1&amp;"x"</f><v>A_x0041_</v></c>""",
            """<c r="G{r}" t="str"><f>"x"</f><v xml:space="preserve"> x </v></c>""",
            """<c r="H{r}" t="b"><f>TRUE()</f><v>1</v></c>""",
            """<c r="I{r}" t="e"><f>1/0</f><v>#DIV/0!</v></c>""",
            """<c r="J{r}" t="s"><f>"s"</f><v>2</v></c>""",
            """<c r="K{r}"><f>NOW()</f></c>""",
            """<c r="L{r}" t="str"><f>A1</f><v/></c>""",
            """<c r="M{r}" t="str"><f>A1</f><v></v></c>""",
            """<c r="N{r}"><f>A1>2</f><v>0</v></c>""",
            """<c r="O{r}"><f/><v>7</v></c>""",
            """<c r="P{r}"><f ca="1">RAND()</f><v>0.25</v></c>""",
            """<c r="Q{r}"><foo/><v>8</v></c>""",
            """<c r="R{r}" s="1" t="d"><f>X</f><v>2024-01-02T03:04:05</v></c>""",
            """<c r="S{r}"><f>A1</f ><v>9</v></c>""",
            """<c t="str"><f>"y"</f><v>no ref</v></c>""",
            """<c r="U{r}" t="s"><f>"s"</f><v>12345678901</v></c>""",
            """<c r="V{r}"><f t="shared" si="0"/></c>""",
        ];

        [Fact]
        public void FormulaCellsReadTheSameOnTheCanonicalPathAsOnTheGeneralPath()
        {
            AssertPathsAgree(FormulaCells, prefixedAgrees: true);
        }

        [Fact]
        public void FormulaCellsReadTheirCachedValues()
        {
            using XlsxWorkbook reader = Excel.FromXlsx(new MemoryStream(Workbook(Rows(FormulaCells, 1))));
            using XlsxWorkbook.Enumerator rows = reader.FirstSheet.GetEnumerator();
            Assert.True(rows.MoveNext());
            Row row = rows.Current;

            Assert.True(row[0].TryGetDouble(out double a));
            Assert.Equal(3, a);
            Assert.Equal("6", row[2].GetString());
            Assert.Equal(CellType.Date, row[3].Type);
            Assert.Equal("a<b", row[4].GetString());
            Assert.Equal("AA", row[5].GetString());
            Assert.Equal(" x ", row[6].GetString());
            Assert.Equal(CellType.Boolean, row[7].Type);
            Assert.Equal("#DIV/0!", row[8].GetString());
            Assert.Equal("gamma", row[9].GetString());
            Assert.Equal("", row[11].GetString());
            Assert.Equal("0", row[13].GetString());
            Assert.Equal("9", row[18].GetString());
            Assert.Equal("no ref", row[19].GetString());
        }

        private static void AssertPathsAgree(string[] cells, bool prefixedAgrees)
        {
            const int RowCount = 300;
            string sheet = Rows(cells, RowCount);
            byte[] plain = Workbook(sheet);
            using MemoryStream general = WorkbookBuilder.Build(sheet.Replace("<c ", "<c cm=\"1\" ", StringComparison.Ordinal), SharedStrings(""), $"<styleSheet>{StylesInner("")}</styleSheet>");

            List<string> expected = Snapshot(general);

            Assert.Equal(RowCount, expected.Count);
            Assert.Equal(expected, Snapshot(new MemoryStream(plain)));
            Assert.Equal(expected, TrickleSnapshot(plain));
            if (!prefixedAgrees)
            {
                return;
            }
            string prefixedSheet = sheet.Replace("<", "<x:", StringComparison.Ordinal).Replace("<x:/", "</x:", StringComparison.Ordinal);
            using MemoryStream prefixed = WorkbookBuilder.BuildPrefixed("x", prefixedSheet, SharedStrings("x:"), StylesInner("x:"));
            Assert.Equal(expected, Snapshot(prefixed));
        }

        private static byte[] Workbook(string sheet)
        {
            using MemoryStream ms = WorkbookBuilder.Build(sheet, SharedStrings(""), $"<styleSheet>{StylesInner("")}</styleSheet>");
            return ms.ToArray();
        }

        private static string Rows(string[] cells, int rows)
        {
            var sb = new StringBuilder();
            for (int r = 1; r <= rows; r++)
            {
                string n = r.ToString(CultureInfo.InvariantCulture);
                sb.Append(CultureInfo.InvariantCulture, $"""<row r="{n}">""");
                foreach (string cell in cells)
                {
                    sb.Append(cell.Replace("{r}", n, StringComparison.Ordinal));
                }
                sb.Append("</row>");
            }
            return sb.ToString();
        }

        // Seven bytes per read, so every row is cut and re-parsed at every few bytes of its cells.
        private static List<string> TrickleSnapshot(byte[] workbook)
        {
            using XlsxWorkbook reader = Excel.FromXlsx(new MemoryStream(workbook));
            reader.OpenSheetWithoutPrefetch(0).Dispose();
            using var sheet = new MemoryStream();
            using (Stream source = reader.OpenSheetStream(0))
            {
                source.CopyTo(sheet);
            }
            reader.Lifetime.Acquire(reader);
            using var rows = new XlsxWorkbook.Enumerator(reader, new TrickleStream(sheet.ToArray()));
            var snapshot = new List<string>();
            while (rows.MoveNext())
            {
                snapshot.Add(Render(rows.Current));
            }
            return snapshot;
        }

        private static string SharedStrings(string p)
        {
            return $"<{p}si><{p}t>alpha</{p}t></{p}si><{p}si><{p}t>beta</{p}t></{p}si><{p}si><{p}t>gamma</{p}t></{p}si>";
        }

        private static string StylesInner(string p)
        {
            return $"""<{p}cellXfs count="2"><{p}xf numFmtId="0"/><{p}xf numFmtId="14" applyNumberFormat="1"/></{p}cellXfs>""";
        }

        private static string SheetRows(string p, int rows)
        {
            var sb = new StringBuilder();
            for (int r = 1; r <= rows; r++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}row r="{r}" spans="1:16">""");
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}c r="A{r}" t="s"><{p}v>{r % 3}</{p}v></{p}c>""");
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}c r="B{r}" s="1"><{p}v>{45000 + r}</{p}v></{p}c>""");
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}c r="C{r}"><{p}v>152.58000000000001</{p}v></{p}c>""");
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}c r="D{r}" t="b"><{p}v>{r % 2}</{p}v></{p}c>""");
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}c r="E{r}" t="e"><{p}v>#N/A</{p}v></{p}c>""");
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}c r="F{r}" t="str"><{p}v>a&amp;b_x0041_</{p}v></{p}c>""");
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}c r="G{r}" s="1"/>""");
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}c t="n"><{p}v>-1.5E-3</{p}v></{p}c>""");
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}c r="J{r}" t="inlineStr"><{p}is><{p}t>inline {r}</{p}t></{p}is></{p}c>""");
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}c r="K{r}" t="s"><{p}v>1234567890</{p}v></{p}c>""");
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}c r="L{r}" s="1" t="d"><{p}v>2024-01-02T03:04:05</{p}v></{p}c>""");
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}c r="M{r}" cm="1"><{p}v>7</{p}v></{p}c>""");
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}c r="N{r}"><{p}f>1+1</{p}f><{p}v>2</{p}v></{p}c>""");
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}c r="O{r}" t="s"><{p}v> 1</{p}v></{p}c>""");
                sb.Append(CultureInfo.InvariantCulture, $"""<{p}c s="1" r="P{r}"><{p}v>{r}.25</{p}v></{p}c>""");
                sb.Append(CultureInfo.InvariantCulture, $"</{p}row>");
            }
            return sb.ToString();
        }

        private static List<string> Snapshot(Stream workbook)
        {
            using XlsxWorkbook reader = Excel.FromXlsx(workbook);
            var rows = new List<string>();
            foreach (Row row in reader.FirstSheet)
            {
                rows.Add(Render(row));
            }
            return rows;
        }

        private static string Render(Row row)
        {
            var sb = new StringBuilder();
            foreach (RowCell rowCell in row.Cells)
            {
                Cell cell = rowCell.Value;
                string number = "-";
                if (cell.TryGetDouble(out double value))
                {
                    number = value.ToString("R", CultureInfo.InvariantCulture);
                }
                string date = cell.TryGetDateTime(false, out DateTime dt) ? dt.ToString("O", CultureInfo.InvariantCulture) : "-";
                sb.Append(CultureInfo.InvariantCulture, $"{rowCell.ColumnIndex}:{cell.Type}:{cell.StyleIndex}:{cell.GetString()}:{number}:{date}|");
            }
            return sb.ToString();
        }
    }
}
