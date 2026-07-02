using System.Globalization;
using System.Text;

namespace ExcelReader.Benchmarks
{
    // Builds self-contained CSV byte[] payloads shaped like WorkbookGenerator's XLSX/XLSB
    // workbooks, so the CSV benchmarks exercise comparable data across formats.
    internal static class CsvGenerator
    {
        // Headerless `rows` of [string, int, date, double] — mirrors WorkbookGenerator.BuildAsync,
        // for the raw cell-by-cell read benchmark. Dates are round-trip ("O") text since CSV has no
        // native date type.
        public static byte[] Build(int rows)
        {
            var sb = new StringBuilder(rows * 40);
            for (int r = 1; r <= rows; r++)
            {
                double serial = 45292 + (r % 3650) + 0.25;
                DateTime date = DateTime.FromOADate(serial);
                sb.Append(WorkbookGenerator.Pool[r % WorkbookGenerator.Pool.Length]).Append(',')
                  .Append(r.ToString(CultureInfo.InvariantCulture)).Append(',')
                  .Append(date.ToString("O", CultureInfo.InvariantCulture)).Append(',')
                  .Append((r * 1.5).ToString(CultureInfo.InvariantCulture)).Append('\n');
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        // Header (Name,Id,Date,Value) + `rows` Records — mirrors WorkbookGenerator.BuildTypedAsync,
        // for the typed-parse benchmark.
        public static byte[] BuildTyped(int rows)
        {
            var sb = new StringBuilder(rows * 40);
            sb.Append("Name,Id,Date,Value\n");
            foreach (Record rec in WorkbookGenerator.Records(rows))
            {
                sb.Append(rec.Name).Append(',')
                  .Append(rec.Id.ToString(CultureInfo.InvariantCulture)).Append(',')
                  .Append(rec.Date.ToString("O", CultureInfo.InvariantCulture)).Append(',')
                  .Append(rec.Value.ToString(CultureInfo.InvariantCulture)).Append('\n');
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }
    }
}
