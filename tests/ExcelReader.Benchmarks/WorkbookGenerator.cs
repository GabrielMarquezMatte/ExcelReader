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

    // Builds self-contained .xlsx workbooks in memory via WorkbookWriter.
    internal static class WorkbookGenerator
    {
        private static readonly string[] Pool =
            ["alpha", "beta", "gamma", "delta", "epsilon", "zeta", "eta", "theta"];

        // `rows` headerless data rows of [string, int, date, float] — exercises
        // every reader value path for the read benchmark.
        public static async Task<byte[]> BuildAsync(int rows)
        {
            await using var ms = new MemoryStream();
            await using (WorkbookWriter wb = await WorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
                SheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync();
                for (int r = 1; r <= rows; r++)
                {
                    double serial = 45292 + (r % 3650) + 0.25; // dates spread over ~10 years
                    await using RowWriter row = await sheet.StartRowAsync();
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
        public static async Task<byte[]> BuildTypedAsync(int rows)
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
                List<Record> records = Records(rows);
                for (int i = 0; i < records.Count; i++)
                {
                    Record rec = records[i];
                    await using RowWriter row = await sheet.StartRowAsync();
                    row.Write(rec.Name);
                    row.Write(rec.Id);
                    row.Write(rec.Date);
                    row.Write(rec.Value);
                }
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.ToArray();
        }
    }
}
