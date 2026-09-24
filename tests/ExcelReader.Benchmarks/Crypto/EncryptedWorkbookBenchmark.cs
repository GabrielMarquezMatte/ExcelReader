using System.Globalization;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;

namespace ExcelReader.Benchmarks
{
    [MemoryDiagnoser]
    public class EncryptedWorkbookBenchmark
    {
        private const string FixturePassword = "hunter2";

        private string _encrypted = "";
        private string _plain = "";
        private byte[] _encryptedBytes = [];

        [GlobalSetup]
        public void Setup()
        {
            string dir = Path.Combine(AppContext.BaseDirectory, "data", "encrypted");
            _encrypted = Environment.GetEnvironmentVariable("EXCELREADER_ENCRYPTED_XLSX")
                ?? Path.Combine(dir, "agile-aes256-sha512.xlsx");
            _plain = Environment.GetEnvironmentVariable("EXCELREADER_PLAIN_XLSX")
                ?? Path.Combine(dir, "agile-aes256-sha512.plain.xlsx");
            _encryptedBytes = File.ReadAllBytes(_encrypted);
        }

        private static ExcelReaderOptions Options(bool verifyIntegrity = false)
        {
            return new ExcelReaderOptions
            {
                Password = new ExcelPassword(FixturePassword),
                VerifyEncryptedIntegrity = verifyIntegrity,
            };
        }

        private static long ReadAll(IExcelRowReader reader)
        {
            long acc = 0;
            using IExcelRowEnumerator rows = reader.GetEnumerator();
            while (rows.MoveNext())
            {
                Row row = rows.Current;
                acc += row.ColumnCount;
            }
            return acc;
        }

        [Benchmark(Baseline = true)]
        public long Plain_Stream()
        {
            using var fs = new FileStream(_plain, FileMode.Open, FileAccess.Read);
            using IExcelRowReader reader = Excel.Open(fs, leaveOpen: true);
            return ReadAll(reader);
        }

        [Benchmark]
        public long Encrypted_Stream()
        {
            using var fs = new FileStream(_encrypted, FileMode.Open, FileAccess.Read);
            using IExcelRowReader reader = Excel.Open(fs, leaveOpen: true, Options());
            return ReadAll(reader);
        }

        [Benchmark]
        public long Encrypted_Stream_VerifyIntegrity()
        {
            using var fs = new FileStream(_encrypted, FileMode.Open, FileAccess.Read);
            using IExcelRowReader reader = Excel.Open(fs, leaveOpen: true, Options(verifyIntegrity: true));
            return ReadAll(reader);
        }

        [Benchmark]
        public long Encrypted_Memory()
        {
            using IExcelRowReader reader = Excel.Open(_encryptedBytes.AsMemory(), Options());
            return ReadAll(reader);
        }

        [Benchmark]
        public int Encrypted_OpenOnly()
        {
            using var ms = new MemoryStream(_encryptedBytes, writable: false);
            using IExcelRowReader reader = Excel.Open(ms, leaveOpen: true, Options());
            return reader.SheetCount;
        }
        private const string GeneratedPassword = "hunter2";
        private const int DefaultLargeRows = 300_000;
        private byte[] _largePlainBytes = [];
        private byte[] _largeEncryptedBytes = [];

        [GlobalSetup(Targets = [nameof(Large_Plain_Stream), nameof(Large_Encrypted_Stream), nameof(Large_Encrypted_OpenOnly)])]
        public async Task LargeSetupAsync()
        {
            int rows = int.TryParse(
                Environment.GetEnvironmentVariable("EXCELREADER_LARGE_ENCRYPTED_ROWS"),
                NumberStyles.Integer, CultureInfo.InvariantCulture, out int configured) && configured > 0
                ? configured
                : DefaultLargeRows;
            _largePlainBytes = await WorkbookGenerator.BuildAsync(rows).ConfigureAwait(false);

            using var plainStream = new MemoryStream(_largePlainBytes, writable: false);
            using var encryptedStream = new MemoryStream();
            Excel.EncryptPackage(plainStream, encryptedStream, new ExcelPassword(GeneratedPassword));
            _largeEncryptedBytes = encryptedStream.ToArray();
        }

        [Benchmark]
        public long Large_Plain_Stream()
        {
            using var ms = new MemoryStream(_largePlainBytes, writable: false);
            using IExcelRowReader reader = Excel.Open(ms, leaveOpen: true);
            return ReadAll(reader);
        }

        [Benchmark]
        public long Large_Encrypted_Stream()
        {
            using var ms = new MemoryStream(_largeEncryptedBytes, writable: false);
            using IExcelRowReader reader = Excel.Open(ms, leaveOpen: true,
                new ExcelReaderOptions { Password = new ExcelPassword(GeneratedPassword) });
            return ReadAll(reader);
        }

        [Benchmark]
        public int Large_Encrypted_OpenOnly()
        {
            using var ms = new MemoryStream(_largeEncryptedBytes, writable: false);
            using IExcelRowReader reader = Excel.Open(ms, leaveOpen: true,
                new ExcelReaderOptions { Password = new(GeneratedPassword) });
            return reader.SheetCount;
        }
    }
}
