using System.Globalization;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xls;
using ExcelReader.Core.Writer.Xls;

namespace ExcelReader.Tests.Reader.Xls
{
    public sealed class XlsLifetimeTests
    {
        private static byte[] BuildXls(int rows = 300)
        {
            using MemoryStream buffer = new();
            using (XlsWorkbookWriter workbook = XlsWorkbookWriter.Create(buffer, leaveOpen: true))
            {
                XlsSheetWriter sheet = workbook.AddSheet("s");
                for (int r = 0; r < rows; r++)
                {
                    using XlsRowWriter row = sheet.StartRow();
                    row.Write("name-" + (r % 40).ToString(CultureInfo.InvariantCulture));
                    row.Write(r);
                }
                sheet.End();
                sheet.Dispose();
                workbook.End();
            }
            return buffer.ToArray();
        }

        private static List<string> Drain(XlsWorkbook.Enumerator e)
        {
            List<string> values = [];
            while (e.MoveNext())
            {
                values.Add(e.Current[0].GetString());
            }
            return values;
        }

        [Fact]
        public void GetEnumerator_After_Dispose_Throws()
        {
            XlsWorkbook reader = Excel.FromXls(BuildXls());
            XlsSheet sheet = reader.FirstSheet;
            reader.Dispose();
            Assert.Throws<ObjectDisposedException>(() => reader.Sheets);
            Assert.Throws<ObjectDisposedException>(() => sheet.GetEnumerator());
            Assert.Throws<ObjectDisposedException>(() => sheet.GetAsyncEnumerator(TestContext.Current.CancellationToken));
        }

        [Fact]
        public void An_Enumerator_Outlives_Its_Reader()
        {
            byte[] bytes = BuildXls();
            List<string> expected;
            using (XlsWorkbook reference = Excel.FromXls(bytes))
            {
                using XlsWorkbook.Enumerator all = reference.FirstSheet.GetEnumerator();
                expected = Drain(all);
            }

            TrickleStream stream = new(bytes);
            XlsWorkbook reader = Excel.FromXls(stream, leaveOpen: false);
            XlsWorkbook.Enumerator e = reader.FirstSheet.GetEnumerator();
            Assert.True(e.MoveNext());

            reader.Dispose();
            Assert.True(stream.CanRead);

            List<string> rest = Drain(e);
            Assert.Equal(expected.Skip(1), rest, StringComparer.Ordinal);

            e.Dispose();
            Assert.False(stream.CanRead);
        }

        [Fact]
        public void Disposing_Reader_And_Enumerator_Twice_Is_Harmless()
        {
            TrickleStream stream = new(BuildXls(rows: 20));
            XlsWorkbook reader = Excel.FromXls(stream, leaveOpen: false);
            XlsWorkbook.Enumerator e = reader.FirstSheet.GetEnumerator();
            e.Dispose();
            e.Dispose();
            Assert.True(stream.CanRead);
            reader.Dispose();
            reader.Dispose();
            Assert.False(stream.CanRead);
        }
    }
}
