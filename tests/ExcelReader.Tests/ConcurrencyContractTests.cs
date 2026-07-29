using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    // Covers the thread-safety contract documented on IExcelRowReader/IWorkbookWriter: instances are
    // not thread-safe, but independent instances used one-per-thread must not interfere with each
    // other. These tests would catch a regression where a reader/writer accidentally leaned on shared
    // mutable state (a static cache, a shared buffer pool misuse, etc.) instead of per-instance state.
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
                using XlsxReader reader = Excel.From(ms);
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
                await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: TestContext.Current.CancellationToken))
                {
                    await wb.StartAsync(TestContext.Current.CancellationToken);
                    XlsxSheetWriter sheet = wb.AddSheet("Sheet1");
                    await sheet.StartAsync(TestContext.Current.CancellationToken);
                    await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                    {
                        row.Write(expected);
                    }
                    await sheet.EndAsync(TestContext.Current.CancellationToken);
                    await wb.EndAsync(TestContext.Current.CancellationToken);
                }
                ms.Position = 0;
                using XlsxReader reader = Excel.From(ms);
                using XlsxReader.Enumerator e = reader.GetEnumerator();
                Assert.True(e.MoveNext());
                Assert.Equal(expected, e.Current[0].GetString());
            }));

            return Task.WhenAll(tasks);
        }
    }
}
