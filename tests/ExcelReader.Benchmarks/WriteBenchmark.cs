using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer;
using OfficeIMO.Excel;
using SpreadCheetah;

namespace ExcelReader.Benchmarks
{
    [MemoryDiagnoser]
    public class WriteBenchmark
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
        public async Task<long> ExcelReaderWriter()
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

        [Benchmark]
        public async Task<long> ExcelReaderWriterSharedStrings()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true, options: new XlsxWriterOptions { UseSharedStrings = true }))
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

        [Benchmark]
        public async Task<long> ExcelReaderWriterPrefetch()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true, options: new XlsxWriterOptions { PrefetchWrite = true }))
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

        [Benchmark]
        public async Task<long> ExcelReaderXlsbWriter()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                ReadOnlySpan<XlsbCell> header =
                [
                    XlsbCell.Create("Name"),
                    XlsbCell.Create("Id"),
                    XlsbCell.Create("Date"),
                    XlsbCell.Create("Value"),
                ];
                sheet.WriteRow(header);
                WorkbookGenerator.WriteXlsbRecords(sheet, _records);
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> ExcelReaderXlsbWriterSharedStrings()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true, options: new XlsbWriterOptions { UseSharedStrings = true }))
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                ReadOnlySpan<XlsbCell> header =
                [
                    XlsbCell.Create("Name"),
                    XlsbCell.Create("Id"),
                    XlsbCell.Create("Date"),
                    XlsbCell.Create("Value"),
                ];
                sheet.WriteRow(header);
                WorkbookGenerator.WriteXlsbRecords(sheet, _records);
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> ExcelReaderXlsbWriterPrefetch()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true, options: new XlsbWriterOptions { PrefetchWrite = true }))
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                ReadOnlySpan<XlsbCell> header =
                [
                    XlsbCell.Create("Name"),
                    XlsbCell.Create("Id"),
                    XlsbCell.Create("Date"),
                    XlsbCell.Create("Value"),
                ];
                sheet.WriteRow(header);
                WorkbookGenerator.WriteXlsbRecords(sheet, _records);
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.Length;
        }

        [Benchmark]
        public long OfficeIMO()
        {
            using var ms = new MemoryStream(4 * 1024 * 1024);
            ExcelDocument.WriteRows(ms, _records, ["Name", "Id", "Date", "Value"], static (row, rec) =>
            {
                row.Write(rec.Name);
                row.Write(rec.Id);
                row.Write(rec.Date);
                row.Write(rec.Value);
            });
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> SpreadCheetah()
        {
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (var writer = await Spreadsheet.CreateNewAsync(ms))
            {
                await writer.StartWorksheetAsync("S1").ConfigureAwait(false);
                await writer.AddHeaderRowAsync(["Name", "Id", "Date", "Value"]).ConfigureAwait(false);
                foreach (var rec in _records)
                {
                    Cell[] row = [
                        new(rec.Name),
                        new(rec.Id),
                        new(rec.Date),
                        new(rec.Value),
                    ];
                    await writer.AddRowAsync(row).ConfigureAwait(false);
                }
                await writer.FinishAsync().ConfigureAwait(false);
            }
            return ms.Length;
        }
    }
}
