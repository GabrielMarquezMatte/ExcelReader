using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.AotSanity
{
    // Exercises the ExcelMappedParser/IExcelRowMap/ExcelRowMapBuilder seam plus the raw, always-AOT-safe
    // Excel.FromCsv reader, under PublishAot=true (see the project file for the AOT/trim settings).
    // Deliberately does not reference the reflection-based typed parser or record writer: doing so must
    // fail "dotnet publish" with a trimming/AOT diagnostic, which is what proves this harness actually
    // detects the problem it exists to catch — confirmed by hand before this file was committed, by
    // temporarily adding such a call here and observing the publish fail.
    internal static class Program
    {
        private static async Task<int> Main()
        {
            await using MemoryStream xlsx = await BuildSampleXlsxAsync();
            await using XlsxReader reader = await Excel.FromAsync(xlsx);
            var rows = new ExcelMappedParser<AotModel>().Parse(reader).ToList();
            if (rows.Count != 1 || !string.Equals(rows[0].Name, "Alice", StringComparison.Ordinal) || rows[0].Age != 30 || !rows[0].Active)
            {
                Console.Error.WriteLine("Mapped XLSX parse produced an unexpected result.");
                return 1;
            }

            ReadOnlyMemory<byte> csv = Encoding.UTF8.GetBytes("Name,Age\r\nBob,42\r\n");
            CsvReader csvReader = Excel.FromCsv(csv);
            CsvReader.Enumerator csvRows = csvReader.GetEnumerator();
            if (!csvRows.MoveNext())
            {
                Console.Error.WriteLine("Raw CSV reader produced no header row.");
                return 1;
            }
            if (!csvRows.MoveNext() || !string.Equals(csvRows.Current[0].GetString(), "Bob", StringComparison.Ordinal))
            {
                Console.Error.WriteLine("Raw CSV reader produced an unexpected data row.");
                return 1;
            }

            Console.WriteLine("AOT sanity checks passed.");
            return 0;
        }

        private static async Task<MemoryStream> BuildSampleXlsxAsync()
        {
            var ms = new MemoryStream();
            await using (XlsxWorkbookWriter wb = await XlsxWorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync();
                await using (XlsxRowWriter header = await sheet.StartRowAsync())
                {
                    header.Write("Name");
                    header.Write("Age");
                    header.Write("Active");
                }
                await using (XlsxRowWriter row = await sheet.StartRowAsync())
                {
                    row.Write("Alice");
                    row.Write(30);
                    row.Write(true);
                }
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            ms.Position = 0;
            return ms;
        }
    }

    internal sealed class AotModel : IExcelRowMap<AotModel>
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
        public bool Active { get; set; }

        public static void ConfigureExcelRowMap(ExcelRowMapBuilder<AotModel> builder)
        {
            builder
                .Factory(static () => new AotModel())
                .Property(["Name"], ExcelCellReaders.String, static (ref AotModel m, string v) => m.Name = v)
                .Property(["Age"], ExcelCellReaders.Parsable<int>, static (ref AotModel m, int v) => m.Age = v)
                .Property(["Active"], ExcelCellReaders.Bool, static (ref AotModel m, bool v) => m.Active = v);
        }
    }
}
