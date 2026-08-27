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

        // 32-column headerless rows, for the CSV read benchmark: the vectorized scan amortizes one
        // vector load over many fields, so a wide row is where the win is largest and a 4-column row
        // (Build above) is where it is smallest. Both are measured so neither shape flatters the result.
        public static byte[] BuildWide(int rows, int columns)
        {
            var sb = new StringBuilder(rows * columns * 6);
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < columns; c++)
                {
                    if (c > 0)
                    {
                        sb.Append(',');
                    }
                    sb.Append(WorkbookGenerator.Pool[(r + c) % WorkbookGenerator.Pool.Length]);
                }
                sb.Append('\n');
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

        // Conversion-heavy: dates, decimals, and several text columns, so per-row conversion
        // dominates the vectorized scan. This is the shape parallelism is supposed to help, and the
        // corpus the 3-5x expectation is stated against.
        internal static string WriteConversionHeavyFile(int rows)
        {
            string path = Path.Combine(Path.GetTempPath(), $"exr-bench-wide-{Guid.NewGuid():N}.csv");
            string[] regions = ["Europe", "Asia", "North America", "Sub-Saharan Africa"];
            string[] countries = ["Portugal", "Japan", "Canada", "Kenya", "Brazil", "Norway"];
            using var writer = new StreamWriter(path, append: false);
            writer.WriteLine("Region,Country,OrderDate,UnitPrice,TotalRevenue,Units");
            var start = new DateTime(2015, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
            for (int i = 0; i < rows; i++)
            {
                DateTime date = start.AddDays(i % 3650);
                decimal price = 10m + (i % 9000 / 100m);
                decimal revenue = price * ((i % 500) + 1);
                writer.Write(regions[i % regions.Length]);
                writer.Write(',');
                writer.Write(countries[i % countries.Length]);
                writer.Write(',');
                writer.Write(date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                writer.Write(',');
                writer.Write(price.ToString(CultureInfo.InvariantCulture));
                writer.Write(',');
                writer.Write(revenue.ToString(CultureInfo.InvariantCulture));
                writer.Write(',');
                writer.WriteLine(((i % 500) + 1).ToString(CultureInfo.InvariantCulture));
            }
            return path;
        }

        // Narrow integers: almost nothing to convert, so the already-vectorized scanner sits near
        // memory bandwidth. This is the shape that shows parallelism's floor, and the reason the
        // benchmark reports two corpora rather than one flattering number.
        internal static string WriteNarrowIntFile(int rows)
        {
            string path = Path.Combine(Path.GetTempPath(), $"exr-bench-narrow-{Guid.NewGuid():N}.csv");
            using var writer = new StreamWriter(path, append: false);
            writer.WriteLine("A,B,C");
            for (int i = 0; i < rows; i++)
            {
                writer.Write(i.ToString(CultureInfo.InvariantCulture));
                writer.Write(',');
                writer.Write((i * 3).ToString(CultureInfo.InvariantCulture));
                writer.Write(',');
                writer.WriteLine((i * 7).ToString(CultureInfo.InvariantCulture));
            }
            return path;
        }
    }
}
