using System.Globalization;
using System.Text;

namespace ExcelReader.Benchmarks
{
    internal static class CsvGenerator
    {
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
