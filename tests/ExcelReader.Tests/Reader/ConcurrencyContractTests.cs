using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Reader
{
    public class ConcurrencyContractTests
    {
        [Fact]
        public Task IndependentXlsxReadersOnSeparateThreadsDoNotInterfere()
        {
            const int readerCount = 16;
            IEnumerable<Task> tasks = Enumerable.Range(0, readerCount).Select(i => Task.Run(() =>
            {
                string expected = $"value-{i}";
                using MemoryStream ms = WorkbookBuilder.Build($"""<row r="1"><c r="A1" t="inlineStr"><is><t>{expected}</t></is></c></row>""");
                using XlsxReader reader = Excel.FromXlsx(ms);
                using XlsxReader.Enumerator e = reader.GetEnumerator();
                Assert.True(e.MoveNext());
                Assert.Equal(expected, e.Current[0].GetString());
            }));

            return Task.WhenAll(tasks);
        }

        [Fact]
        public Task IndependentXlsxWorkbookWritersOnSeparateThreadsDoNotInterfere()
        {
            const int writerCount = 16;
            IEnumerable<Task> tasks = Enumerable.Range(0, writerCount).Select(i => Task.Run(async () =>
            {
                string expected = $"value-{i}";
                using MemoryStream ms = new();
                await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
                {
                    XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
                    await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                    {
                        row.Write(expected);
                    }
                    await sheet.EndAsync(TestContext.Current.CancellationToken);
                    await wb.EndAsync(TestContext.Current.CancellationToken);
                }
                ms.Position = 0;
                using XlsxReader reader = Excel.FromXlsx(ms);
                using XlsxReader.Enumerator e = reader.GetEnumerator();
                Assert.True(e.MoveNext());
                Assert.Equal(expected, e.Current[0].GetString());
            }));

            return Task.WhenAll(tasks);
        }
    }
}
