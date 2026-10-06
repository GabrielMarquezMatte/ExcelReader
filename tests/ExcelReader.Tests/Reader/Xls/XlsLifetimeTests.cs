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

        private static List<string> Drain(XlsReader.Enumerator e)
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
            XlsReader reader = Excel.FromXls(BuildXls());
            reader.Dispose();
            Assert.Throws<ObjectDisposedException>(() => reader.GetEnumerator());
            Assert.Throws<ObjectDisposedException>(() => reader.GetAsyncEnumerator(TestContext.Current.CancellationToken));
        }

        [Fact]
        public void An_Enumerator_Outlives_Its_Reader()
        {
            byte[] bytes = BuildXls();
            List<string> expected;
            using (XlsReader reference = Excel.FromXls(bytes))
            {
                using XlsReader.Enumerator all = reference.GetEnumerator();
                expected = Drain(all);
            }

            TrickleStream stream = new(bytes);
            XlsReader reader = Excel.FromXls(stream, leaveOpen: false);
            XlsReader.Enumerator e = reader.GetEnumerator();
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
            XlsReader reader = Excel.FromXls(stream, leaveOpen: false);
            XlsReader.Enumerator e = reader.GetEnumerator();
            e.Dispose();
            e.Dispose();
            Assert.True(stream.CanRead);
            reader.Dispose();
            reader.Dispose();
            Assert.False(stream.CanRead);
        }
    }
}
