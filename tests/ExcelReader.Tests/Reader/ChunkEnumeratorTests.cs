using System.Globalization;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer.Xlsb;

namespace ExcelReader.Tests.Reader
{
    public class ChunkEnumeratorTests
    {
        private static string Render(Row row)
        {
            var sb = new StringBuilder();
            foreach (RowCell cell in row.Cells)
            {
                sb.Append(cell.ColumnIndex).Append(':').Append(cell.Value.GetString()).Append('|');
            }
            return sb.ToString();
        }

        private static string Rows(int count)
        {
            var sb = new StringBuilder();
            for (int r = 1; r <= count; r++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"<row r=\"{r}\"><c r=\"A{r}\" t=\"inlineStr\"><is><t>n{r}</t></is></c><c r=\"B{r}\"><v>{r}</v></c></row>");
            }
            return sb.ToString();
        }

        private static byte[] Inflated(XlsxWorkbook workbook)
        {
            using Stream stream = workbook.OpenSheetStream(0);
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            return copy.ToArray();
        }

        [Fact]
        public void AnXlsxChunkParsesLikeTheSameRowsReadSequentially()
        {
            using MemoryStream ms = WorkbookBuilder.Build(Rows(50));
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            var expected = new List<string>();
            long dataStart;
            NsTokens? ns;
            using (XlsxWorkbook.Enumerator rows = workbook.OpenSheetWithoutPrefetch(0))
            {
                ns = rows.Namespace;
                Assert.True(rows.MoveNext());
                dataStart = rows.Offset;
                while (rows.MoveNext())
                {
                    expected.Add(Render(rows.Current));
                }
            }
            byte[] inflated = Inflated(workbook);

            var actual = new List<string>();
            using var chunk = new XlsxWorkbook.Enumerator(workbook, inflated.AsMemory((int)dataStart), ns, CancellationToken.None);
            while (chunk.MoveNext())
            {
                actual.Add(Render(chunk.Current));
            }

            Assert.Equal(expected, actual);
            Assert.True(chunk.ReachedEnd);
            Assert.False(chunk.EndedMidMarkup);
        }

        [Fact]
        public void AnXlsxChunkCutBetweenRowsEndsCleanly()
        {
            using MemoryStream ms = WorkbookBuilder.Build(Rows(10));
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            using (workbook.OpenSheetWithoutPrefetch(0))
            {
            }
            byte[] inflated = Inflated(workbook);
            int first = inflated.AsSpan().IndexOf("<row r=\"1\""u8);
            int sixth = inflated.AsSpan().IndexOf("<row r=\"6\""u8);

            using var chunk = new XlsxWorkbook.Enumerator(workbook, inflated.AsMemory(first, sixth - first), null, CancellationToken.None);
            int rows = 0;
            while (chunk.MoveNext())
            {
                rows++;
            }

            Assert.Equal(5, rows);
            Assert.False(chunk.ReachedEnd);
            Assert.False(chunk.EndedMidMarkup);
        }

        [Fact]
        public void AnXlsxChunkCutInsideARowReportsIt()
        {
            using MemoryStream ms = WorkbookBuilder.Build(Rows(10));
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            using (workbook.OpenSheetWithoutPrefetch(0))
            {
            }
            byte[] inflated = Inflated(workbook);
            int first = inflated.AsSpan().IndexOf("<row r=\"1\""u8);
            int insideSixth = inflated.AsSpan().IndexOf("<row r=\"6\""u8) + 20;

            using var chunk = new XlsxWorkbook.Enumerator(workbook, inflated.AsMemory(first, insideSixth - first), null, CancellationToken.None);
            int complete = 0;
            while (chunk.MoveNext() && !chunk.EndedMidMarkup)
            {
                complete++;
            }

            Assert.Equal(5, complete);
            Assert.True(chunk.EndedMidMarkup);
        }

        [Fact]
        public void APrefixedXlsxSheetReportsItsNamespaceTokens()
        {
            using MemoryStream ms = WorkbookBuilder.BuildPrefixed("x", "<x:row r=\"1\"><x:c r=\"A1\"><x:v>1</x:v></x:c></x:row>");
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            using XlsxWorkbook.Enumerator rows = workbook.OpenSheetWithoutPrefetch(0);

            Assert.NotNull(rows.Namespace);
            Assert.Equal("<x:row"u8.ToArray(), rows.Namespace.RowOpen);
        }

        [Fact]
        public async Task AnXlsbChunkResumesInsideTheRowTheHeaderReadLeftOpen()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var ms = new MemoryStream();
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                for (int r = 0; r < 30; r++)
                {
                    await using XlsbRowWriter row = await sheet.StartRowAsync(ct);
                    row.Write("n" + r.ToString(CultureInfo.InvariantCulture));
                    row.Write(r);
                }
                await sheet.EndAsync(ct);
                await wb.EndAsync(ct);
            }
            ms.Position = 0;
            using XlsbWorkbook workbook = Excel.FromXlsb(ms);
            var expected = new List<string>();
            long dataStart;
            bool rowOpen;
            using (XlsbWorkbook.Enumerator rows = workbook.OpenSheetWithoutPrefetch(0))
            {
                Assert.True(rows.MoveNext());
                dataStart = rows.Offset;
                rowOpen = rows.RowOpen;
                while (rows.MoveNext())
                {
                    expected.Add(Render(rows.Current));
                }
            }
            byte[] inflated;
            using (Stream stream = workbook.OpenSheetStream(0))
            using (var copy = new MemoryStream())
            {
                await stream.CopyToAsync(copy, ct);
                inflated = copy.ToArray();
            }

            var actual = new List<string>();
            using var chunk = new XlsbWorkbook.Enumerator(workbook, inflated.AsMemory((int)dataStart), rowOpen, CancellationToken.None);
            while (chunk.MoveNext())
            {
                actual.Add(Render(chunk.Current));
            }

            Assert.True(rowOpen);
            Assert.Equal(29, expected.Count);
            Assert.Equal(expected, actual);
        }
    }
}
