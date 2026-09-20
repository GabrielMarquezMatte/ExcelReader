using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Core.Writer;

namespace ExcelReader.AotSanity
{
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

            var generatedRows = new ExcelMappedParser<GeneratedAotModel>().Parse(reader).ToList();
            if (generatedRows.Count != 1 || !string.Equals(generatedRows[0].Name, "Alice", StringComparison.Ordinal)
                || generatedRows[0].Age != 30 || !generatedRows[0].Active)
            {
                Console.Error.WriteLine("Source-generated XLSX parse produced an unexpected result.");
                return 1;
            }

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

            if (!EncryptedWorkbookReadsUnderAot())
            {
                return 1;
            }

            Console.WriteLine("AOT sanity checks passed.");
            return 0;
        }

        private static bool EncryptedWorkbookReadsUnderAot()
        {
            string encrypted = Path.Combine(AppContext.BaseDirectory, "data", "encrypted", "agile-aes256-sha512.xlsx");
            var options = new ExcelReaderOptions { Password = "hunter2" };
            using IExcelRowReader encryptedReader = Excel.Open(encrypted, options);
            int encryptedRowCount = 0;
            foreach (Row encryptedRow in encryptedReader)
            {
                encryptedRowCount++;
            }
            if (encryptedRowCount == 0)
            {
                Console.Error.WriteLine("FAIL: encrypted workbook yielded no rows under AOT");
                return false;
            }
            Console.WriteLine($"OK: encrypted workbook read {encryptedRowCount} rows under AOT");
            return true;
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
