using ExcelReader.Core.Enums;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Core.Writer;

namespace ExcelReader.CodSpeed
{
    // The measured bodies. Each method is one pass over the 65K-row dataset (read/parse) or one full
    // serialization of the synthetic record list (write). Every value is folded into a checksum so the
    // JIT cannot elide the work being measured.
    //
    // The read loops are written out per reader type on purpose: iterating through IExcelRowReader
    // would dispatch every row through an interface and measure that indirection instead of the
    // concrete enumerator an application actually uses.
    internal static class Workloads
    {
        public const int WriteRows = 50_000;

        private const int WriteBufferBytes = 16 * 1024 * 1024;

        private static readonly ExcelReaderOptions PrefetchOptions = new() { PrefetchDecompression = true };

        // --- readers ---

        public static long ReadXlsxStream()
        {
            using MemoryStream ms = new(Fixtures.Xlsx, writable: false);
            using XlsxReader reader = Excel.From(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        // True in-memory path: no MemoryStream, the ZIP central directory is read straight out of the
        // caller's buffer.
        public static long ReadXlsxMemory()
        {
            using XlsxReader reader = Excel.From(Fixtures.Xlsx.AsMemory());
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        public static long ReadXlsxPrefetch()
        {
            using MemoryStream ms = new(Fixtures.Xlsx, writable: false);
            using XlsxReader reader = Excel.From(ms, options: PrefetchOptions);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        // Materializes a managed string per text cell instead of reading the zero-copy UTF-8 span:
        // the allocating path an application that keeps the values pays for.
        public static long ReadXlsxMaterialized()
        {
            using MemoryStream ms = new(Fixtures.Xlsx, writable: false);
            using XlsxReader reader = Excel.From(ms);
            long acc = 0;
            foreach (Row row in reader)
            {
                foreach (RowCell rowCell in row.Cells)
                {
                    Cell cell = rowCell.Value;
                    switch (cell.Type)
                    {
                        case CellType.ExcelString:
                            acc += cell.GetString().Length;
                            break;
                        case CellType.Number:
                            if (cell.TryParse(null, out double n)) { acc += (long)n; }
                            break;
                        case CellType.Date:
                            if (cell.TryGetDateTime(out DateTime d)) { acc += d.Ticks; }
                            break;
                        default:
                            break;
                    }
                }
            }
            return acc;
        }

        public static async Task<long> ReadXlsxAsync()
        {
            await using MemoryStream ms = new(Fixtures.Xlsx, writable: false);
            await using XlsxReader reader = await Excel.FromAsync(ms);
            await using XlsxReader.Enumerator rows = await reader.GetAsyncEnumeratorAsync();
            long acc = 0;
            while (await rows.MoveNextAsync())
            {
                acc += AccumulateRow(rows.Current);
            }
            return acc;
        }

        public static long ReadXlsbStream()
        {
            using MemoryStream ms = new(Fixtures.Xlsb, writable: false);
            using XlsbReader reader = Excel.FromXlsb(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        public static long ReadXlsStream()
        {
            using MemoryStream ms = new(Fixtures.Xls, writable: false);
            using XlsReader reader = Excel.FromXls(ms);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        public static long ReadCsvStream()
        {
            using MemoryStream ms = new(Fixtures.Csv, writable: false);
            using CsvReader reader = Excel.FromCsv(ms);
            long acc = 0;
            foreach (Row row in reader)
            {
                foreach (RowCell rowCell in row.Cells)
                {
                    acc += rowCell.Value.Value.Length;
                }
            }
            return acc;
        }

        // --- typed parsing ---

        public static long ParseXlsxClass()
        {
            using MemoryStream ms = new(Fixtures.Xlsx, writable: false);
            using XlsxReader reader = Excel.From(ms);
            long acc = 0;
            foreach (SalesRecord record in new ExcelParser<SalesRecord>().Parse(reader))
            {
                acc += Accumulate(record);
            }
            return acc;
        }

        // Same columns as ParseXlsxClass, but the model is a struct: ExcelParser<T> binds through
        // `ref TModel`, so this path never allocates a model per row.
        public static long ParseXlsxStruct()
        {
            using MemoryStream ms = new(Fixtures.Xlsx, writable: false);
            using XlsxReader reader = Excel.From(ms);
            long acc = 0;
            foreach (SalesRecordStruct record in new ExcelParser<SalesRecordStruct>().Parse(reader))
            {
                acc += Accumulate(record);
            }
            return acc;
        }

        public static long ParseCsvClass()
        {
            using MemoryStream ms = new(Fixtures.Csv, writable: false);
            using CsvReader reader = Excel.FromCsv(ms);
            long acc = 0;
            foreach (SalesRecord record in new ExcelParser<SalesRecord>().Parse(reader))
            {
                acc += Accumulate(record);
            }
            return acc;
        }

        // --- writers ---

        public static async Task<long> WriteXlsxAsync(List<WriteRecord> records, bool useSharedStrings)
        {
            await using MemoryStream ms = new(WriteBufferBytes);
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true, useSharedStrings: useSharedStrings))
            {
                await wb.StartAsync();
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync();
                using (XlsxRowWriter header = sheet.StartRow()) { WriteHeader(header); }
                for (int i = 0; i < records.Count; i++)
                {
                    WriteRecord record = records[i];
                    using XlsxRowWriter row = sheet.StartRow();
                    row.Write(record.Name);
                    row.Write(record.Id);
                    row.Write(record.Date);
                    row.Write(record.Value);
                }
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.Length;
        }

        public static async Task<long> WriteXlsbAsync(List<WriteRecord> records)
        {
            await using MemoryStream ms = new(WriteBufferBytes);
            await using (XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync();
                await using (XlsbRowWriter header = await sheet.StartRowAsync()) { WriteHeader(header); }
                for (int i = 0; i < records.Count; i++)
                {
                    WriteRecord record = records[i];
                    await using XlsbRowWriter row = await sheet.StartRowAsync();
                    row.Write(record.Name);
                    row.Write(record.Id);
                    row.Write(record.Date);
                    row.Write(record.Value);
                }
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.Length;
        }

        public static async Task<long> WriteXlsAsync(List<WriteRecord> records)
        {
            await using MemoryStream ms = new(WriteBufferBytes);
            await using (XlsWorkbookWriter wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                wb.Start();
                XlsSheetWriter sheet = wb.AddSheet("S1");
                sheet.Start();
                using (XlsRowWriter header = sheet.StartRow()) { WriteHeader(header); }
                for (int i = 0; i < records.Count; i++)
                {
                    WriteRecord record = records[i];
                    using XlsRowWriter row = sheet.StartRow();
                    row.Write(record.Name);
                    row.Write(record.Id);
                    row.Write(record.Date);
                    row.Write(record.Value);
                }
                sheet.End();
                await wb.EndAsync();
            }
            return ms.Length;
        }

        public static long WriteCsv(List<WriteRecord> records)
        {
            using MemoryStream ms = new(WriteBufferBytes);
            using (CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true))
            {
                using (CsvRowWriter header = writer.StartRow()) { WriteHeader(header); }
                for (int i = 0; i < records.Count; i++)
                {
                    WriteRecord record = records[i];
                    using CsvRowWriter row = writer.StartRow();
                    row.Write(record.Name);
                    row.Write(record.Id);
                    row.Write(record.Date);
                    row.Write(record.Value);
                }
            }
            return ms.Length;
        }

        // High-level POCO dump: the API most applications reach for, which routes every property
        // through the generic Write<T> overload.
        public static async Task<long> WriteRecordsXlsxAsync(List<WriteRecord> records)
        {
            await using MemoryStream ms = new(WriteBufferBytes);
            await using (var writer = await RecordWriter.CreateXlsxAsync(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", records);
            }
            return ms.Length;
        }

        // --- helpers ---

        private static void WriteHeader(IRowWriter header)
        {
            header.Write("Name");
            header.Write("Id");
            header.Write("Date");
            header.Write("Value");
        }

        private static long AccumulateRow(Row row)
        {
            long acc = 0;
            foreach (RowCell rowCell in row.Cells)
            {
                Cell cell = rowCell.Value;
                switch (cell.Type)
                {
                    case CellType.ExcelString:
                        acc += cell.Value.Length;
                        break;
                    case CellType.Number:
                        if (cell.TryParse(null, out double n)) { acc += (long)n; }
                        break;
                    case CellType.Date:
                        if (cell.TryGetDateTime(out DateTime d)) { acc += d.Ticks; }
                        break;
                    default:
                        break;
                }
            }
            return acc;
        }

        private static long Accumulate(SalesRecord record)
        {
            return (record.Region?.Length ?? 0)
                + (record.ItemType?.Length ?? 0)
                + record.OrderId
                + record.UnitsSold
                + (long)record.TotalProfit
                + record.OrderDate.Ticks;
        }

        private static long Accumulate(SalesRecordStruct record)
        {
            return (record.Region?.Length ?? 0)
                + (record.ItemType?.Length ?? 0)
                + record.OrderId
                + record.UnitsSold
                + (long)record.TotalProfit
                + record.OrderDate.Ticks;
        }
    }
}
