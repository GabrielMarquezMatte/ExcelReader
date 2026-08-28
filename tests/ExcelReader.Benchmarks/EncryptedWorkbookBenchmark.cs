using System.Diagnostics.CodeAnalysis;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Benchmarks
{
    // Reading a password-protected (ECMA-376 agile) workbook, against the same workbook unencrypted.
    // The pair is the point: the plain leg is what reading this file costs anyway, so the difference
    // is what encryption actually adds.
    //
    // The corpus is the checked-in oracle fixture pair from the test suite (encrypted with
    // `hunter2`, AES-256 / SHA-512 / spinCount 100,000). It is small on purpose — writing encrypted
    // workbooks is out of scope for this library, so there is no way to generate a large one here
    // without depending on msoffcrypto-tool at benchmark time. That makes this suite a measurement of
    // the *fixed* cost of opening an encrypted workbook, which is dominated by the spinCount key
    // derivation and does not shrink with file size; the per-byte AES cost is a rounding error at
    // this size and is not what these numbers are about.
    //
    // Point EXCELREADER_ENCRYPTED_XLSX / EXCELREADER_PLAIN_XLSX at a larger pair to measure the
    // per-byte side instead. The password still has to be `hunter2`.
    [MemoryDiagnoser]
    public class EncryptedWorkbookBenchmark
    {
        // Not a credential: this is the published password of the repository's own encrypted test
        // fixtures, documented in tests/ExcelReader.Tests/data/encrypted/README.md, guarding files
        // that contain nothing but generated benchmark rows.
        [SuppressMessage("Major Code Smell", "S2068:Hard-coded credentials are security-sensitive",
            Justification = "Published password of the repository's own public test fixtures; there is no secret here to leak.")]
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

        // What the same workbook costs with no encryption in the way.
        [Benchmark(Baseline = true)]
        public long Plain_Stream()
        {
            using var fs = new FileStream(_plain, FileMode.Open, FileAccess.Read);
            using IExcelRowReader reader = Excel.Open(fs, leaveOpen: true);
            return ReadAll(reader);
        }

        // The streaming path: decrypts 4 KB segments on demand as ZipArchive reads them.
        [Benchmark]
        public long Encrypted_Stream()
        {
            using var fs = new FileStream(_encrypted, FileMode.Open, FileAccess.Read);
            using IExcelRowReader reader = Excel.Open(fs, leaveOpen: true, Options());
            return ReadAll(reader);
        }

        // Same, plus the opt-in dataIntegrity HMAC, which is a full extra pass over the ciphertext
        // before the first row is produced.
        [Benchmark]
        public long Encrypted_Stream_VerifyIntegrity()
        {
            using var fs = new FileStream(_encrypted, FileMode.Open, FileAccess.Read);
            using IExcelRowReader reader = Excel.Open(fs, leaveOpen: true, Options(verifyIntegrity: true));
            return ReadAll(reader);
        }

        // The ReadOnlyMemory<byte> overload, which is documented never to suspend and therefore
        // decrypts the whole package eagerly (and always verifies its integrity).
        [Benchmark]
        public long Encrypted_Memory()
        {
            using IExcelRowReader reader = Excel.Open(_encryptedBytes.AsMemory(), Options());
            return ReadAll(reader);
        }

        // Open without reading a single row: isolates the fixed cost — CFB parse, descriptor parse,
        // and the spinCount password-key derivation — from anything that scales with the package.
        [Benchmark]
        public int Encrypted_OpenOnly()
        {
            using var ms = new MemoryStream(_encryptedBytes, writable: false);
            using IExcelRowReader reader = Excel.Open(ms, leaveOpen: true, Options());
            return reader.SheetCount;
        }
    }
}
