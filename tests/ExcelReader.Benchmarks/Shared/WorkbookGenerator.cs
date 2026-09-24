using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Benchmarks
{
    public sealed class Record
    {
        public string? Name { get; set; }
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public double Value { get; set; }
    }

    public struct RecordStruct
    {
        public string? Name { get; set; }
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public double Value { get; set; }
    }

    public readonly ref struct RecordNamedRef
    {
        public ReadOnlySpan<byte> Name { get; init; }
        public int Id { get; init; }
        public DateTime Date { get; init; }
        public double Value { get; init; }
    }

    internal static class WorkbookGenerator
    {
        internal static readonly string[] Pool =
            ["alpha", "beta", "gamma", "delta", "epsilon", "zeta", "eta", "theta"];

        public static Task<byte[]> BuildAsync(int rows)
        {
            return BuildAsync<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(rows, static ms => XlsxWorkbookWriter.Create(ms, leaveOpen: true));
        }

        public static Task<byte[]> BuildXlsbAsync(int rows)
        {
            return BuildAsync<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(rows, static ms => XlsbWorkbookWriter.Create(ms, leaveOpen: true));
        }

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

        public static async Task<byte[]> BuildTypedAsync(int rows)
        {
            await using var ms = new MemoryStream();
            await using (var writer = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync(Records(rows), ExcelRecordLayout.FromAttributes<Record>());
                }
            }
            return ms.ToArray();
        }

        public static async Task<byte[]> BuildTypedSharedStringsAsync(int rows)
        {
            await using var ms = new MemoryStream();
            await using (var writer = XlsxWorkbookWriter.Create(ms, leaveOpen: true, options: new XlsxWriterOptions { UseSharedStrings = true }))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync(Records(rows), ExcelRecordLayout.FromAttributes<Record>());
                }
            }
            return ms.ToArray();
        }

        public static async Task<byte[]> BuildTypedXlsbAsync(int rows)
        {
            await using var ms = new MemoryStream();
            await using (var writer = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await sheet.WriteRecordsAsync(Records(rows), ExcelRecordLayout.FromAttributes<Record>());
                }
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
            Func<MemoryStream, TWorkbook> create)
            where TWorkbook : IWorkbookWriter<TSheet>
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            await using var ms = new MemoryStream();
            await using (TWorkbook wb = create(ms))
            {
                TSheet sheet = wb.AddSheet("S1");
                for (int r = 1; r <= rows; r++)
                {
                    double serial = 45292 + (r % 3650) + 0.25;
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
