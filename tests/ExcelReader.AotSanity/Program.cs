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

            // Same check, but through the source-generated IExcelRowMap<T> (ExcelSerializableAttribute)
            // instead of the hand-written one above — proves the generator's own emitted code, not just
            // the hand-written seam, survives PublishAot.
            var generatedRows = new ExcelMappedParser<GeneratedAotModel>().Parse(reader).ToList();
            if (generatedRows.Count != 1 || !string.Equals(generatedRows[0].Name, "Alice", StringComparison.Ordinal)
                || generatedRows[0].Age != 30 || !generatedRows[0].Active)
            {
                Console.Error.WriteLine("Source-generated XLSX parse produced an unexpected result.");
                return 1;
            }

            // Feature A4: the public MappedRecordWriter.CreateMapped*Async write entries, driven by the
            // same source-generated map (IExcelRecordMap<T>), also under PublishAot.
            await using var writtenStream = new MemoryStream();
            await using (var writer = await MappedRecordWriter.CreateMappedXlsxAsync(writtenStream, leaveOpen: true))
            {
                await writer.WriteSheetAsync("S1", [new GeneratedAotModel { Name = "Zoe", Age = 8, Active = true }]);
            }
            writtenStream.Position = 0;
            await using XlsxReader writtenReader = await Excel.FromAsync(writtenStream);
            var writtenRows = new ExcelMappedParser<GeneratedAotModel>().Parse(writtenReader).ToList();
            if (writtenRows.Count != 1 || !string.Equals(writtenRows[0].Name, "Zoe", StringComparison.Ordinal) || writtenRows[0].Age != 8 || !writtenRows[0].Active)
            {
                Console.Error.WriteLine("Source-generated XLSX write+read round trip produced an unexpected result.");
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
                .Property(["Name"], ExcelCellReaders.String, static (ref m, v) => m.Name = v)
                .Property(["Age"], ExcelCellReaders.Parsable, static (ref AotModel m, int v) => m.Age = v)
                .Property(["Active"], ExcelCellReaders.Bool, static (ref m, v) => m.Active = v);
        }
    }

    [ExcelSerializable]
    internal sealed partial class GeneratedAotModel
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
        public bool Active { get; set; }
    }
}
