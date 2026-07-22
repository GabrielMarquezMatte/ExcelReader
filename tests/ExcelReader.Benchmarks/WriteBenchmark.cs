using System.Diagnostics.CodeAnalysis;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer;
using MiniExcelLibs;
using SpreadCheetah;

namespace ExcelReader.Benchmarks
{
    // Writes `Rows` records (header + 4 columns) to an in-memory .xlsx, comparing
    // ExcelReader's XlsxWorkbookWriter against MiniExcel's object serializer.
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
        [SuppressMessage("Sonar", "S6966:Await StartRowAsync instead",
            Justification = "Deliberately measuring XlsxSheetWriter's synchronous row fast path (StartRow/XlsxRowWriter.Dispose).")]
        [SuppressMessage("VisualStudio.Threading", "VSTHRD103:StartRow synchronously blocks",
            Justification = "Deliberately measuring XlsxSheetWriter's synchronous row fast path (StartRow/XlsxRowWriter.Dispose).")]
        public async Task<long> ExcelReaderWriter()
        {
            // Pre-sized to the neighborhood of the actual output so MemoryStream's doubling growth
            // (256B -> ... -> 4MB) doesn't dominate the GC/allocation numbers being measured.
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync();
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
        [SuppressMessage("Sonar", "S6966:Await StartRowAsync instead",
            Justification = "Deliberately measuring XlsxSheetWriter's synchronous row fast path (StartRow/XlsxRowWriter.Dispose).")]
        [SuppressMessage("VisualStudio.Threading", "VSTHRD103:StartRow synchronously blocks",
            Justification = "Deliberately measuring XlsxSheetWriter's synchronous row fast path (StartRow/XlsxRowWriter.Dispose).")]
        public async Task<long> ExcelReaderWriterSharedStrings()
        {
            // Pre-sized to the neighborhood of the actual output so MemoryStream's doubling growth
            // (256B -> ... -> 4MB) doesn't dominate the GC/allocation numbers being measured.
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true, useSharedStrings: true))
            {
                await wb.StartAsync();
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync();
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
            // Pre-sized to the neighborhood of the actual output so MemoryStream's doubling growth
            // (256B -> ... -> 4MB) doesn't dominate the GC/allocation numbers being measured.
            await using var ms = new MemoryStream(4 * 1024 * 1024);
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
        public async Task<long> ExcelReaderXlsbWriterSharedStrings()
        {
            // Pre-sized to the neighborhood of the actual output so MemoryStream's doubling growth
            // (256B -> ... -> 4MB) doesn't dominate the GC/allocation numbers being measured.
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await using (XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, useSharedStrings: true))
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
            // Pre-sized to the neighborhood of the actual output so MemoryStream's doubling growth
            // (256B -> ... -> 4MB) doesn't dominate the GC/allocation numbers being measured.
            await using var ms = new MemoryStream(4 * 1024 * 1024);
            await ms.SaveAsAsync(_records, excelType: ExcelType.XLSX).ConfigureAwait(false);
            return ms.Length;
        }

        [Benchmark]
        public async Task<long> SpreadCheetah()
        {
            // Pre-sized to the neighborhood of the actual output so MemoryStream's doubling growth
            // (256B -> ... -> 4MB) doesn't dominate the GC/allocation numbers being measured.
            await using var ms = new MemoryStream(4 * 1024 * 1024);
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
