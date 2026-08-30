using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    public class EncryptedIntegrityTests
    {
        private static ExcelReaderOptions WithPassword =>
            ExcelReaderOptions.Default with { Password = EncryptedFixtures.Password };

        private static int CountRows(IExcelRowReader reader)
        {
            int n = 0;
            foreach (Row _ in reader)
            {
                n++;
            }
            return n;
        }

        // The memory overload documents that it never suspends, which forces eager decryption -
        // so verification there is nearly free and always on.
        [Fact]
        public void Should_Read_When_Opened_From_Memory_With_Password()
        {
            byte[] bytes = EncryptedFixtures.Bytes("agile-aes256-sha512.xlsx");
            using XlsxReader reader = Excel.From(bytes, WithPassword);
            Assert.True(CountRows(reader) > 0);
        }

        [Fact]
        public void Should_Report_PasswordRequired_When_Memory_Path_Has_No_Password()
        {
            byte[] bytes = EncryptedFixtures.Bytes("agile-aes256-sha512.xlsx");
            var ex = Assert.Throws<ExcelEncryptionException>(() => Excel.Open(bytes));
            Assert.Equal(ExcelEncryptionReason.PasswordRequired, ex.Reason);
        }

        // Tampering must be reported as tampering, not as a wrong password - otherwise a user with
        // a corrupt file retries passwords forever.
        [Fact]
        public void Should_Report_IntegrityFailure_When_Ciphertext_Tampered_On_Memory_Path()
        {
            byte[] bytes = EncryptedFixtures.Bytes("agile-aes256-sha512.xlsx");
            int offset = FindCiphertextOffset(bytes);
            bytes[offset] ^= 0xFF;

            var ex = Assert.Throws<ExcelEncryptionException>(() => Excel.Open(bytes, WithPassword));
            Assert.Equal(ExcelEncryptionReason.IntegrityFailure, ex.Reason);
        }

        // Default off on the streaming path: verifying would need a full pass before the first row.
        [Fact]
        public void Should_Not_Verify_By_Default_On_Streaming_Path()
        {
            Assert.False(ExcelReaderOptions.Default.VerifyEncryptedIntegrity);
            using FileStream fs = File.OpenRead(EncryptedFixtures.Path_("agile-aes256-sha512.xlsx"));
            using IExcelRowReader reader = Excel.Open(fs, leaveOpen: true, WithPassword);
            Assert.True(CountRows(reader) > 0);
        }

        [Fact]
        public void Should_Verify_When_Opted_In_On_Streaming_Path()
        {
            var opted = WithPassword with { VerifyEncryptedIntegrity = true };
            using FileStream fs = File.OpenRead(EncryptedFixtures.Path_("agile-aes256-sha512.xlsx"));
            using IExcelRowReader reader = Excel.Open(fs, leaveOpen: true, opted);
            Assert.True(CountRows(reader) > 0);
        }

        [Fact]
        public void Should_Report_IntegrityFailure_When_Opted_In_And_Tampered()
        {
            byte[] bytes = EncryptedFixtures.Bytes("agile-aes256-sha512.xlsx");
            bytes[FindCiphertextOffset(bytes)] ^= 0xFF;
            string temp = Path.Combine(Path.GetTempPath(), $"tampered-{Guid.NewGuid():N}.xlsx");
            File.WriteAllBytes(temp, bytes);
            try
            {
                var opted = WithPassword with { VerifyEncryptedIntegrity = true };
                var ex = Assert.Throws<ExcelEncryptionException>(() => Excel.Open(temp, opted));
                Assert.Equal(ExcelEncryptionReason.IntegrityFailure, ex.Reason);
            }
            finally
            {
                File.Delete(temp);
            }
        }

        // Not every agile file has a dataIntegrity element (older writers can omit it); opting in
        // must be a no-op then too, not a spurious failure. Simulated by stripping the element from
        // a real descriptor rather than a standard-encryption fixture, since none exists in this
        // pass (see "Execution Scope Note") — the wiring rule ("only AgileDescriptor, only when
        // dataIntegrity is present") is the same thing a standard-encryption fixture would exercise.
        [Fact]
        public void Should_Open_When_Opted_In_But_Descriptor_Has_No_DataIntegrity()
        {
            byte[] bytes = EncryptedFixtures.Bytes("agile-aes256-sha512.xlsx");
            string containerText = System.Text.Encoding.Latin1.GetString(bytes);
            int start = containerText.IndexOf("<dataIntegrity", StringComparison.Ordinal);
            int end = containerText.IndexOf("/>", start, StringComparison.Ordinal) + 2;
            // Overwrite the element with spaces rather than removing it: keeps every offset in the
            // OLE container - and every other field in EncryptionInfo - unchanged.
            Array.Fill(bytes, (byte)' ', start, end - start);
            string temp = Path.Combine(Path.GetTempPath(), $"no-integrity-{Guid.NewGuid():N}.xlsx");
            File.WriteAllBytes(temp, bytes);
            try
            {
                var opted = WithPassword with { VerifyEncryptedIntegrity = true };
                using IExcelRowReader reader = Excel.Open(temp, opted);
                Assert.True(CountRows(reader) > 0);
            }
            finally
            {
                File.Delete(temp);
            }
        }

        // Flips a byte deep inside the EncryptedPackage ciphertext. The stream is stored in whole
        // 512-byte CFB sectors, and the fixture is multi-segment, so an offset 3/4 of the way into
        // the file lands in ciphertext rather than in container metadata.
        private static int FindCiphertextOffset(byte[] container)
        {
            return container.Length * 3 / 4;
        }
    }
}
