using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer;
using MiniExcelLibs;
using SpreadCheetah;

namespace ExcelReader.Benchmarks
{
    // Writes `Rows` records (header + 4 columns) to an in-memory .xlsx, comparing
    // ExcelReader's WorkbookWriter against MiniExcel's object serializer.
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
            await using var ms = new MemoryStream();
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


        [Benchmark]
        public async Task<long> ExcelReaderXlsbWriter()
        {
            await using var ms = new MemoryStream();
            await using (XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync();
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
        public async Task<long> MiniExcel()
        {
            await using var ms = new MemoryStream();
            await ms.SaveAsAsync(_records, excelType: ExcelType.XLSX).ConfigureAwait(false);
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> SpreadCheetah()
        {
            await using var ms = new MemoryStream();
            await using (var writer = await Spreadsheet.CreateNewAsync(ms))
            {
                await writer.StartWorksheetAsync("S1").ConfigureAwait(false);
                await writer.AddHeaderRowAsync(["Name", "Id", "Date", "Value"]).ConfigureAwait(false);
                foreach(var rec in _records)
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
