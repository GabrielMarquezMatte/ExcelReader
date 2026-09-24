using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Benchmarks
{
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
            await using var ms = new MemoryStream(16 * 1024 * 1024);
            await using (XlsWorkbookWriter wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsSheetWriter sheet = wb.AddSheet("S1");
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
            await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                using (XlsxRowWriter header = sheet.StartRow())
                {
                    header.Write("Name");
                    header.Write("Id");
                    header.Write("Date");
                    header.Write("Value");
                }
                for (int i = 0; i < _records.Count; i++)
                {
                    Record rec = _records[i];
                    using XlsxRowWriter row = sheet.StartRow();
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
