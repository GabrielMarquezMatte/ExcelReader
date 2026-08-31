using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public sealed class EncryptedPackageWriteTests
    {
        private const string Password = "hunter2";

        private static List<string[]> ReadAllRows(IExcelRowReader reader)
        {
            List<string[]> rows = [];
            using IExcelRowEnumerator e = reader.GetEnumerator();
            while (e.MoveNext())
            {
                List<string> cells = [];
                foreach (var cell in e.Current.Cells)
                {
                    cells.Add(cell.Value.GetString());
                }
                rows.Add([.. cells]);
            }
            return rows;
        }

        private static byte[] Encrypt(string plainFixture)
        {
            using FileStream plain = File.OpenRead(plainFixture);
            using var encrypted = new MemoryStream();
            Excel.EncryptPackage(plain, encrypted, Password);
            return encrypted.ToArray();
        }

        [Theory]
        [InlineData("agile-aes256-sha512.xlsx")]
        [InlineData("agile-aes256-sha512.xlsb")]
        public void EncryptPackage_RoundTrips_Through_The_Reader(string fixture)
        {
            string plainPath = EncryptedFixtures.PlainPath(fixture);
            byte[] encrypted = Encrypt(plainPath);

            List<string[]> expected;
            using (IExcelRowReader plainReader = Excel.Open(plainPath))
            {
                expected = ReadAllRows(plainReader);
            }

            ExcelReaderOptions options = new() { Password = Password, VerifyEncryptedIntegrity = true };
            using IExcelRowReader decrypted = Excel.Open(encrypted, options);

            Assert.Equal(expected.Count, ReadAllRows(decrypted).Count);
        }

        [Fact]
        public void EncryptPackage_Output_Requires_A_Password_To_Open()
        {
            byte[] encrypted = Encrypt(EncryptedFixtures.PlainPath("agile-aes256-sha512.xlsx"));

            ExcelEncryptionException exception = Assert.Throws<ExcelEncryptionException>(
                () => Excel.Open(encrypted, ExcelReaderOptions.Default));
            Assert.Equal(ExcelEncryptionReason.PasswordRequired, exception.Reason);
        }

        [Fact]
        public void EncryptPackage_Output_Rejects_The_Wrong_Password()
        {
            byte[] encrypted = Encrypt(EncryptedFixtures.PlainPath("agile-aes256-sha512.xlsx"));

            ExcelReaderOptions options = new() { Password = "not-the-password" };
            ExcelEncryptionException exception = Assert.Throws<ExcelEncryptionException>(
                () => Excel.Open(encrypted, options));
            Assert.Equal(ExcelEncryptionReason.PasswordIncorrect, exception.Reason);
        }

        [Fact]
        public void EncryptPackage_Stores_EncryptionInfo_In_The_Mini_Stream()
        {
            byte[] encrypted = Encrypt(EncryptedFixtures.PlainPath("agile-aes256-sha512.xlsx"));

            using var container = new MemoryStream(encrypted, writable: false);
            using CfbContainer cfb = CfbContainer.Parse(container, ownsSource: false, ExcelReaderOptions.Default);
            long infoSize = cfb.StreamLength("EncryptionInfo");

            Assert.True(infoSize > 0 && infoSize < 4096, $"EncryptionInfo was {infoSize} bytes.");
            Assert.NotEmpty(cfb.ReadStream("EncryptionInfo", 64 * 1024));
        }

        [Fact]
        public async Task EncryptPackageAsync_PathOverload_RoundTrips()
        {
            string plainPath = EncryptedFixtures.PlainPath("agile-aes256-sha512.xlsx");
            string outputPath = Path.Combine(Path.GetTempPath(), $"excelreader-encrypt-{Guid.NewGuid():N}.xlsx");
            try
            {
                await Excel.EncryptPackageAsync(plainPath, outputPath, Password, TestContext.Current.CancellationToken);

                ExcelReaderOptions options = new() { Password = Password, VerifyEncryptedIntegrity = true };
                using IExcelRowReader reader = Excel.Open(outputPath, options);
                Assert.NotEmpty(ReadAllRows(reader));
            }
            finally
            {
                File.Delete(outputPath);
            }
        }

        [Fact]
        public void EncryptPackage_NullArguments_Throw()
        {
            using var stream = new MemoryStream([0x50, 0x4B, 0x03, 0x04], writable: false);
            Assert.Throws<ArgumentNullException>(() => Excel.EncryptPackage(null!, stream, Password));
            Assert.Throws<ArgumentNullException>(() => Excel.EncryptPackage(stream, null!, Password));
            Assert.Throws<ArgumentNullException>(() => Excel.EncryptPackage(stream, stream, null!));
        }
    }
}
