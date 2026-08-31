using System.Security.Cryptography;
using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public sealed class PackageEncryptorTests
    {
        private const string Password = "hunter2";

        // A plaintext that is deliberately not a multiple of either 4096 or 16, so the final short
        // segment and its padding are exercised.
        private static byte[] SamplePackage(int length = 10_000)
        {
            byte[] content = new byte[length];
            RandomNumberGenerator.Fill(content);
            content[0] = (byte)'P';
            content[1] = (byte)'K';
            content[2] = 0x03;
            content[3] = 0x04;
            return content;
        }

        [Fact]
        public void Encrypt_Produces_A_Container_Holding_Both_Streams()
        {
            byte[] plain = SamplePackage();
            using var source = new MemoryStream(plain, writable: false);
            using var destination = new MemoryStream();

            PackageEncryptor.Encrypt(source, destination, Password);

            using var container = new MemoryStream(destination.ToArray(), writable: false);
            using CfbContainer cfb = CfbContainer.Parse(container, ownsSource: false, ExcelReaderOptions.Default);
            Assert.True(cfb.ContainsStream("EncryptionInfo"));
            Assert.True(cfb.ContainsStream("EncryptedPackage"));
            // 8-byte prefix + ciphertext rounded up to the cipher block.
            Assert.Equal(8 + (((plain.Length + 15) / 16) * 16), cfb.StreamLength("EncryptedPackage"));
        }

        [Fact]
        public void Encrypt_Output_Decrypts_Back_To_The_Original_Bytes()
        {
            byte[] plain = SamplePackage();
            using var source = new MemoryStream(plain, writable: false);
            using var destination = new MemoryStream();
            PackageEncryptor.Encrypt(source, destination, Password);

            ExcelReaderOptions options = new() { Password = Password, VerifyEncryptedIntegrity = true };
            ReadOnlyMemory<byte> decrypted = EncryptedPackageOpener.DecryptToMemory(destination.ToArray(), options);

            Assert.Equal(plain, decrypted.ToArray());
        }

        [Fact]
        public void Encrypt_TamperedCiphertext_Fails_The_Integrity_Check()
        {
            using var source = new MemoryStream(SamplePackage(), writable: false);
            using var destination = new MemoryStream();
            PackageEncryptor.Encrypt(source, destination, Password);

            byte[] container = destination.ToArray();
            container[TamperableOffset(container, "EncryptedPackage")] ^= 0xFF;

            ExcelReaderOptions options = new() { Password = Password, VerifyEncryptedIntegrity = true };
            ExcelEncryptionException exception = Assert.Throws<ExcelEncryptionException>(
                () => EncryptedPackageOpener.DecryptToMemory(container, options));
            Assert.Equal(ExcelEncryptionReason.IntegrityFailure, exception.Reason);
        }

        [Fact]
        public async Task EncryptAsync_Output_Decrypts_Back_To_The_Original_Bytes()
        {
            byte[] plain = SamplePackage(5_000);
            using var source = new MemoryStream(plain, writable: false);
            using var destination = new MemoryStream();

            await PackageEncryptor.EncryptAsync(source, destination, Password, TestContext.Current.CancellationToken);

            ExcelReaderOptions options = new() { Password = Password, VerifyEncryptedIntegrity = true };
            ReadOnlyMemory<byte> decrypted = EncryptedPackageOpener.DecryptToMemory(destination.ToArray(), options);
            Assert.Equal(plain, decrypted.ToArray());
        }

        [Fact]
        public void Encrypt_MultiSegmentPackage_DecryptsBack()
        {
            // > 2 segments, so a wrong per-segment IV or segment counter shows up here and nowhere else.
            byte[] plain = SamplePackage(9_500);
            using var source = new MemoryStream(plain, writable: false);
            using var destination = new MemoryStream();
            PackageEncryptor.Encrypt(source, destination, Password);

            ExcelReaderOptions options = new() { Password = Password, VerifyEncryptedIntegrity = true };
            Assert.Equal(plain, EncryptedPackageOpener.DecryptToMemory(destination.ToArray(), options).ToArray());
        }

        [Fact]
        public void Encrypt_NonOpcInput_Throws()
        {
            using var source = new MemoryStream(new byte[512], writable: false);
            using var destination = new MemoryStream();
            Assert.Throws<ArgumentException>(() => PackageEncryptor.Encrypt(source, destination, Password));
        }

        [Fact]
        public void Encrypt_EmptyPackage_Throws()
        {
            using var source = new MemoryStream([], writable: false);
            using var destination = new MemoryStream();
            Assert.Throws<ArgumentException>(() => PackageEncryptor.Encrypt(source, destination, Password));
        }

        [Fact]
        public void Encrypt_PackageShorterThanSignature_ThrowsArgumentException()
        {
            using var source = new MemoryStream([0x50, 0x4B], writable: false);
            using var destination = new MemoryStream();
            Assert.Throws<ArgumentException>(() => PackageEncryptor.Encrypt(source, destination, Password));
        }

        [Fact]
        public void Encrypt_NonWritableDestination_ThrowsArgumentException()
        {
            using var source = new MemoryStream(SamplePackage(), writable: false);
            using var destination = new MemoryStream([], writable: false);
            Assert.Throws<ArgumentException>(() => PackageEncryptor.Encrypt(source, destination, Password));
        }

        // The index of a stream's last real data byte, as opposed to the zero padding a big stream
        // carries out to its final sector boundary — nothing reads that padding, so flipping a byte
        // in it cannot be detected as tampering.
        private static int TamperableOffset(byte[] container, string name)
        {
            using var view = new MemoryStream(container, writable: false);
            using CfbContainer cfb = CfbContainer.Parse(view, ownsSource: false, ExcelReaderOptions.Default);
            Assert.True(cfb.TryFindEntry(name, out CfbContainer.DirectoryEntry entry));
            long offset = ((entry.StartSector + 1L) * cfb.SectorSize) + entry.Size - 1;
            return checked((int)offset);
        }
    }
}
