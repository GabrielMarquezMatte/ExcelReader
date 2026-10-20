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
                var sb = new StringBuilder();
                foreach (RowCell rowCell in row.Cells)
                {
                    Cell cell = rowCell.Value;
                    string number = "-";
                    if (cell.TryGetDouble(out double value))
                    {
                        number = value.ToString("R", CultureInfo.InvariantCulture);
                    }
                    sb.Append(CultureInfo.InvariantCulture, $"{rowCell.ColumnIndex}:{cell.Type}:{cell.StyleIndex}:{cell.GetString()}:{number}|");
                }
                rows.Add(sb.ToString());
            }
            return rows;
        }
    }
}
