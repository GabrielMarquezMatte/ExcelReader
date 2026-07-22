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

    // Struct twin of Record, same property names/types. ExcelParser<T> binds properties via `ref
    // TModel` throughout (ColumnParser<T>/RefAction<T,TProperty> — see Delegates.cs), so parsing into
    // a struct T avoids the per-row model allocation that a class T requires. Used by
    // ParseBenchmark.ExcelParserStructSync to measure that directly.
    public struct RecordStruct
    {
        public string? Name { get; set; }
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public double Value { get; set; }
    }

    // ref struct twin of Record, parsed via RefParser.ParseNamed<T> (reflection/attribute-driven —
    // ExcelReader.Core.Parser.RefParser) rather than IExcelRowModel<T>.FromRow. A genuine `ref struct`
    // (not just a normal struct, which ExcelParser<T> already supports) — proves ParseNamed's
    // reflection pipeline works for ref structs too. Name stays `string?` (allocates per row): see
    // RefParser.ParseNamed's doc comment — span-typed property binding isn't implemented yet.
    public readonly ref struct RecordNamedRef
    {
        public ReadOnlySpan<byte> Name { get; init; }
        public int Id { get; init; }
        public DateTime Date { get; init; }
        public double Value { get; init; }
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
            return BuildAsync<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(rows, static ms => XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true));
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

        // Workbook with a header row (Name/Id/Date/Value) + `rows` data rows, for the typed-parse
        // benchmark. Built through the high-level WorkbookRecordWriter (header from property names +
        // one row per record), the same path an application would use to emit typed data.
        public static async Task<byte[]> BuildTypedAsync(int rows)
        {
            await using var ms = new MemoryStream();
            await using (var writer = await RecordWriter.CreateXlsxAsync(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", Records(rows));
            }
            return ms.ToArray();
        }

        public static async Task<byte[]> BuildTypedXlsbAsync(int rows)
        {
            await using var ms = new MemoryStream();
            await using (var writer = await RecordWriter.CreateXlsbAsync(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", Records(rows));
            }
            return ms.ToArray();
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
    }
}
