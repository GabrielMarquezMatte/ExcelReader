using System.Security.Cryptography;
using System.Text;
using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public sealed class EncryptionInfoWriterTests
    {
        private static CryptoParameters Params(byte[] salt, int spinCount, byte[] verifierInput, byte[] verifierValue, byte[] keyValue)
        {
            return new CryptoParameters(
                SaltSize: 16, BlockSize: 16, KeyBits: 256, HashSize: 64, Hash: HashKind.Sha512,
                SaltValue: salt, SpinCount: spinCount,
                EncryptedVerifierHashInput: verifierInput,
                EncryptedVerifierHashValue: verifierValue,
                EncryptedKeyValue: keyValue);
        }

        private static byte[] BuildSample(out CryptoParameters keyData, out CryptoParameters passwordEncryptor,
            out byte[] hmacKey, out byte[] hmacValue)
        {
            keyData = Params(RandomNumberGenerator.GetBytes(16), 0, [], [], []);
            passwordEncryptor = Params(RandomNumberGenerator.GetBytes(16), 100_000,
                RandomNumberGenerator.GetBytes(16), RandomNumberGenerator.GetBytes(64), RandomNumberGenerator.GetBytes(32));
            hmacKey = RandomNumberGenerator.GetBytes(64);
            hmacValue = RandomNumberGenerator.GetBytes(64);
            return EncryptionInfoWriter.Build(keyData, passwordEncryptor, hmacKey, hmacValue);
        }

        [Fact]
        public void Build_Output_Parses_Back_To_The_Same_Parameters()
        {
            byte[] info = BuildSample(out CryptoParameters keyData, out CryptoParameters passwordEncryptor,
                out byte[] hmacKey, out byte[] hmacValue);

            var parsed = (AgileDescriptor)EncryptionDescriptor.Parse(info, ExcelReaderOptions.Default);

            Assert.Equal(keyData.SaltValue, parsed.KeyData.SaltValue);
            Assert.Equal(256, parsed.KeyData.KeyBits);
            Assert.Equal(16, parsed.KeyData.BlockSize);
            Assert.Equal(64, parsed.KeyData.HashSize);
            Assert.Equal(HashKind.Sha512, parsed.KeyData.Hash);
            Assert.Equal(passwordEncryptor.SaltValue, parsed.PasswordEncryptor.SaltValue);
            Assert.Equal(100_000, parsed.PasswordEncryptor.SpinCount);
            Assert.Equal(passwordEncryptor.EncryptedVerifierHashInput, parsed.PasswordEncryptor.EncryptedVerifierHashInput);
            Assert.Equal(passwordEncryptor.EncryptedVerifierHashValue, parsed.PasswordEncryptor.EncryptedVerifierHashValue);
            Assert.Equal(passwordEncryptor.EncryptedKeyValue, parsed.PasswordEncryptor.EncryptedKeyValue);
            Assert.Equal(hmacKey, parsed.EncryptedHmacKey);
            Assert.Equal(hmacValue, parsed.EncryptedHmacValue);
            Assert.True(parsed.HasDataIntegrity);
        }

        [Fact]
        public void Build_Emits_The_Version_Header_And_Xml_Prologue_Excel_Writes()
        {
            byte[] info = BuildSample(out _, out _, out _, out _);

            Assert.Equal<byte[]>([0x04, 0x00, 0x04, 0x00, 0x40, 0x00, 0x00, 0x00], info[..8]);
            Assert.StartsWith(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<encryption ",
                Encoding.UTF8.GetString(info, 8, info.Length - 8),
                StringComparison.Ordinal);
        }

        // Compares our descriptor's shape against a real Excel-written one, parsed through the same
        // parser: same cipher parameters, same spin count, dataIntegrity present.
        [Fact]
        public void Build_Parameters_Match_A_Real_Excel_Descriptor()
        {
            byte[] ours = BuildSample(out _, out _, out _, out _);
            var mine = (AgileDescriptor)EncryptionDescriptor.Parse(ours, ExcelReaderOptions.Default);

            using FileStream fixture = File.OpenRead(EncryptedFixtures.Path_("agile-aes256-sha512.xlsx"));
            using CfbContainer cfb = CfbContainer.Parse(fixture, ownsSource: false, ExcelReaderOptions.Default);
            var excel = (AgileDescriptor)EncryptionDescriptor.Parse(
                cfb.ReadStream("EncryptionInfo", 64 * 1024), ExcelReaderOptions.Default);

            Assert.Equal(excel.KeyData.KeyBits, mine.KeyData.KeyBits);
            Assert.Equal(excel.KeyData.BlockSize, mine.KeyData.BlockSize);
            Assert.Equal(excel.KeyData.HashSize, mine.KeyData.HashSize);
            Assert.Equal(excel.KeyData.SaltSize, mine.KeyData.SaltSize);
            Assert.Equal(excel.KeyData.Hash, mine.KeyData.Hash);
            Assert.Equal(excel.PasswordEncryptor.SpinCount, mine.PasswordEncryptor.SpinCount);
            Assert.Equal(excel.PasswordEncryptor.KeyBits, mine.PasswordEncryptor.KeyBits);
            Assert.Equal(excel.PasswordEncryptor.SaltSize, mine.PasswordEncryptor.SaltSize);
            Assert.Equal(excel.PasswordEncryptor.Hash, mine.PasswordEncryptor.Hash);
            Assert.True(excel.HasDataIntegrity && mine.HasDataIntegrity);
        }
    }
}
