using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer;

namespace ExcelReader.Benchmarks
{
    // Writes `Rows` records (header + 4 columns) to an in-memory .xls via XlsWorkbookWriter.
    // No third-party legacy-.xls writer exists in the referenced packages, so the .xlsx writer
    // (same data) is included as a same-codebase reference point: binary BIFF8 vs zipped XML.
    [MemoryDiagnoser]
    public class XlsWriteBenchmark
    {
        [Params(50_000)]
        public int Rows { get; set; }

        private List<Record> _records = [];

        [GlobalSetup]
        public void Setup()
        {
            _records = WorkbookGenerator.Records(Rows);
        }

        [Benchmark(Baseline = true)]
        public async Task<long> XlsWriter()
        {
            // Pre-sized to the neighborhood of the actual output so MemoryStream's doubling growth
            // doesn't dominate the GC/allocation numbers being measured.
            await using var ms = new MemoryStream(16 * 1024 * 1024);
            await using (XlsWorkbookWriter wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                wb.Start();
                XlsSheetWriter sheet = wb.AddSheet("S1");
                sheet.Start();
                using (XlsRowWriter header = sheet.StartRow())
                {
                    header.Write("Name");
                    header.Write("Id");
                    header.Write("Date");
                    header.Write("Value");
                }
                for (int i = 0; i < _records.Count; i++)
                {
                    Record rec = _records[i];
                    using XlsRowWriter row = sheet.StartRow();
                    row.Write(rec.Name);
                    row.Write(rec.Id);
                    row.Write(rec.Date);
                    row.Write(rec.Value);
                }
                sheet.End();
                await wb.EndAsync();
            }
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> XlsxWriter()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (WorkbookWriter wb = await WorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
                SheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync();
                await using (RowWriter header = await sheet.StartRowAsync())
                {
                    header.Write("Name");
                    header.Write("Id");
                    header.Write("Date");
                    header.Write("Value");
                }
                for (int i = 0; i < _records.Count; i++)
                {
                    Record rec = _records[i];
                    await using RowWriter row = await sheet.StartRowAsync();
                    row.Write(rec.Name);
                    row.Write(rec.Id);
                    row.Write(rec.Date);
                    row.Write(rec.Value);
                }
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.Length;
        }
    }
}
