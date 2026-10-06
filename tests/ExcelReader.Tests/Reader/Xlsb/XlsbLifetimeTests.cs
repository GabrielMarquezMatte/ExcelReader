using System.Globalization;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Writer.Xlsb;

namespace ExcelReader.Tests.Reader.Xlsb
{
    public sealed class XlsbLifetimeTests
    {
        private static byte[] BuildXlsb(int rows = 300)
        {
            using MemoryStream buffer = new();
            using (XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(buffer, leaveOpen: true, new XlsbWriterOptions { UseSharedStrings = true }))
            {
                XlsbSheetWriter sheet = workbook.AddSheet("s");
                for (int r = 0; r < rows; r++)
                {
                    using XlsbRowWriter row = sheet.StartRow();
                    row.Write("name-" + (r % 40).ToString(CultureInfo.InvariantCulture));
                    row.Write(r);
                }
                sheet.End();
                sheet.Dispose();
                workbook.End();
            }
            return buffer.ToArray();
        }

        private static List<string> Drain(XlsbWorkbook.Enumerator e)
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
            XlsbWorkbook reader = Excel.FromXlsb(BuildXlsb());
            XlsbSheet sheet = reader.FirstSheet;
            reader.Dispose();
            Assert.Throws<ObjectDisposedException>(() => reader.Sheets);
            Assert.Throws<ObjectDisposedException>(() => sheet.GetEnumerator());
            Assert.Throws<ObjectDisposedException>(() => sheet.GetAsyncEnumerator(TestContext.Current.CancellationToken));
        }

        [Fact]
        public void An_Enumerator_Outlives_Its_Reader()
        {
            byte[] bytes = BuildXlsb();
            List<string> expected;
            using (XlsbWorkbook reference = Excel.FromXlsb(bytes))
            {
                using XlsbWorkbook.Enumerator all = reference.FirstSheet.GetEnumerator();
                expected = Drain(all);
            }

            TrickleStream stream = new(bytes);
            XlsbWorkbook reader = Excel.FromXlsb(stream, leaveOpen: false);
            XlsbWorkbook.Enumerator e = reader.FirstSheet.GetEnumerator();
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
            TrickleStream stream = new(BuildXlsb(rows: 20));
            XlsbWorkbook reader = Excel.FromXlsb(stream, leaveOpen: false);
            XlsbWorkbook.Enumerator e = reader.FirstSheet.GetEnumerator();
            e.Dispose();
            e.Dispose();
            Assert.True(stream.CanRead);
            reader.Dispose();
            reader.Dispose();
            Assert.False(stream.CanRead);
        }
    }
}
