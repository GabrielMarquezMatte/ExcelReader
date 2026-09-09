using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Benchmarks
{
    // Reading a password-protected (ECMA-376 agile) workbook, against the same workbook unencrypted.
    // The pair is the point: the plain leg is what reading this file costs anyway, so the difference
    // is what encryption actually adds.
    //
    // Two corpora, two different things measured:
    //
    // - Encrypted_*/Plain_Stream (below) use the checked-in oracle fixture pair from the test suite
    //   (encrypted with `hunter2`, AES-256 / SHA-512 / spinCount 100,000). It is small on purpose, so
    //   what these legs measure is the *fixed* cost of opening an encrypted workbook — dominated by
    //   the spinCount key derivation, independent of file size. The per-byte AES cost is a rounding
    //   error at this size and is not what these numbers are about. Point
    //   EXCELREADER_ENCRYPTED_XLSX / EXCELREADER_PLAIN_XLSX at a larger pair to override this corpus
    //   directly (the password still has to be `hunter2`), or use the Large_* legs below instead.
    // - Large_* legs generate their own large pair in-process (see LargeSetup's comment) specifically
    //   to measure the per-byte side the small fixture can't show.
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
        [SuppressMessage("Major Code Smell", "S2068:Hard-coded credentials are security-sensitive",
            Justification = "Password for an in-memory workbook this benchmark generates and discards itself.")]
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

        // What the generated large workbook costs with no encryption in the way — the baseline the
        // two legs below are measured against.
        [Benchmark]
        public long Large_Plain_Stream()
        {
            using var ms = new MemoryStream(_largePlainBytes, writable: false);
            using IExcelRowReader reader = Excel.Open(ms, leaveOpen: true);
            return ReadAll(reader);
        }

        // The per-byte cost this whole section exists to measure: decrypting and reading a package
        // far larger than the fixed-cost-dominated legs above can show.
        [Benchmark]
        public long Large_Encrypted_Stream()
        {
            using var ms = new MemoryStream(_largeEncryptedBytes, writable: false);
            using IExcelRowReader reader = Excel.Open(ms, leaveOpen: true,
                new ExcelReaderOptions { Password = new ExcelPassword(GeneratedPassword) });
            return ReadAll(reader);
        }

        // Same fixed-cost isolation as Encrypted_OpenOnly, at the large package's size — confirms the
        // open-only cost doesn't grow with the package (it shouldn't: CFB directory parse and
        // spinCount derivation are both independent of EncryptedPackage's length).
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
