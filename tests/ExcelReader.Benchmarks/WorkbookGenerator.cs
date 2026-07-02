using ExcelReader.Core.Writer;

namespace ExcelReader.Benchmarks
{
    // Strongly-typed row used by the write and parse benchmarks.
    // Property names match the header written by BuildTypedAsync.
    public sealed class Record
    {
        public string? Name { get; set; }
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public double Value { get; set; }
    }

    // Builds self-contained workbooks in memory via the project writers.
    internal static class WorkbookGenerator
    {
        internal static readonly string[] Pool =
            ["alpha", "beta", "gamma", "delta", "epsilon", "zeta", "eta", "theta"];

        // `rows` headerless data rows of [string, int, date, float] — exercises
        // every reader value path for the read benchmark.
        public static Task<byte[]> BuildAsync(int rows)
        {
            return BuildAsync<WorkbookWriter, SheetWriter, RowWriter>(rows, static ms => WorkbookWriter.CreateAsync(ms, leaveOpen: true));
        }

        public static Task<byte[]> BuildXlsbAsync(int rows)
        {
            return BuildAsync<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(rows, static ms => XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true));
        }

        // `rows` Records — source data for write benchmarks.
        public static List<Record> Records(int rows)
        {
            var list = new List<Record>(rows);
            for (int r = 1; r <= rows; r++)
            {
                list.Add(new Record
                {
                    Name = Pool[r % Pool.Length],
                    Id = r,
                    Date = DateTime.FromOADate(45292 + (r % 3650) + 0.25),
                    Value = r * 1.5,
                });
            }
            return list;
        }

        // Workbook with a header row (Name/Id/Date/Value) + `rows` data rows,
        // for the typed-parse benchmark.
        public static Task<byte[]> BuildTypedAsync(int rows)
        {
            return BuildTypedAsync<WorkbookWriter, SheetWriter, RowWriter>(rows, static ms => WorkbookWriter.CreateAsync(ms, leaveOpen: true));
        }

        public static Task<byte[]> BuildTypedXlsbAsync(int rows)
        {
            return BuildTypedAsync<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(rows, static ms => XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true));
        }

        internal static async Task WriteRecordsAsync<TRow>(ISheetWriter<TRow> sheet, List<Record> records)
            where TRow : IRowWriter
        {
            for (int i = 0; i < records.Count; i++)
            {
                Record rec = records[i];
                await using TRow row = await sheet.StartRowAsync();
                row.Write(rec.Name);
                row.Write(rec.Id);
                row.Write(rec.Date);
                row.Write(rec.Value);
            }
        }

        internal static void WriteXlsbRecords(XlsbSheetWriter sheet, List<Record> records)
        {
            XlsbCell[] row = new XlsbCell[4];
            for (int i = 0; i < records.Count; i++)
            {
                Record rec = records[i];
                row[0] = XlsbCell.Create(rec.Name);
                row[1] = XlsbCell.Create(rec.Id);
                row[2] = XlsbCell.Create(rec.Date);
                row[3] = XlsbCell.Create(rec.Value);
                sheet.WriteRow(row);
            }
        }

        private static async Task<byte[]> BuildAsync<TWorkbook, TSheet, TRow>(
            int rows,
            Func<MemoryStream, ValueTask<TWorkbook>> create)
            where TWorkbook : IWorkbookWriter<TSheet>
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            await using var ms = new MemoryStream();
            await using (TWorkbook wb = await create(ms))
            {
                await wb.StartAsync();
                TSheet sheet = wb.AddSheet("S1");
                await sheet.StartAsync();
                for (int r = 1; r <= rows; r++)
                {
                    double serial = 45292 + (r % 3650) + 0.25; // dates spread over ~10 years
                    await using TRow row = await sheet.StartRowAsync();
                    row.Write(Pool[r % Pool.Length]);
                    row.Write(r);
                    row.Write(DateTime.FromOADate(serial));
                    row.Write(r * 1.5);
                }
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.ToArray();
        }

        private static async Task<byte[]> BuildTypedAsync<TWorkbook, TSheet, TRow>(
            int rows,
            Func<MemoryStream, ValueTask<TWorkbook>> create)
            where TWorkbook : IWorkbookWriter<TSheet>
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            await using var ms = new MemoryStream();
            await using (TWorkbook wb = await create(ms))
            {
                await wb.StartAsync();
                TSheet sheet = wb.AddSheet("S1");
                await sheet.StartAsync();
                await using (TRow header = await sheet.StartRowAsync())
                {
                    header.Write("Name");
                    header.Write("Id");
                    header.Write("Date");
                    header.Write("Value");
                }
                await WriteRecordsAsync(sheet, Records(rows));
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.ToArray();
        }
    }
}
