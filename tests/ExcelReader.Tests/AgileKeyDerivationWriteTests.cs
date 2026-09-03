using System.Security.Cryptography;
using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public sealed class AgileKeyDerivationWriteTests
    {
        private const string Password = "hunter2";

        private static CryptoParameters PasswordEncryptor(byte[] salt, PasswordEncryptorBlobs blobs)
        {
            return new CryptoParameters(
                SaltSize: 16, BlockSize: 16, KeyBits: 256, HashSize: 64, Hash: HashKind.Sha512,
                SaltValue: salt, SpinCount: 100_000,
                EncryptedVerifierHashInput: blobs.EncryptedVerifierHashInput,
                EncryptedVerifierHashValue: blobs.EncryptedVerifierHashValue,
                EncryptedKeyValue: blobs.EncryptedKeyValue);
        }

        private static CryptoParameters KeyData(byte[] salt)
        {
            return new CryptoParameters(
                SaltSize: 16, BlockSize: 16, KeyBits: 256, HashSize: 64, Hash: HashKind.Sha512,
                SaltValue: salt, SpinCount: 0,
                EncryptedVerifierHashInput: [], EncryptedVerifierHashValue: [], EncryptedKeyValue: []);
        }

        [Fact]
        public void DeriveWriteBlobs_Output_Is_Readable_By_DeriveIntermediateKey()
        {
            byte[] pwSalt = RandomNumberGenerator.GetBytes(16);
            byte[] keySalt = RandomNumberGenerator.GetBytes(16);
            byte[] packageKey = RandomNumberGenerator.GetBytes(32);
            byte[] verifierInput = RandomNumberGenerator.GetBytes(16);

            CryptoParameters seed = PasswordEncryptor(pwSalt, default);
            PasswordEncryptorBlobs blobs = AgileKeyDerivation.DeriveWriteBlobs(seed, Password, packageKey, verifierInput);
            AgileDescriptor descriptor = new(KeyData(keySalt), PasswordEncryptor(pwSalt, blobs), [], []);

            byte[] recovered = AgileKeyDerivation.DeriveIntermediateKey(descriptor, Password);

            Assert.Equal(packageKey, recovered);
        }

        [Fact]
        public void DeriveIntermediateKey_WithWrongPassword_Throws_PasswordIncorrect()
        {
            byte[] pwSalt = RandomNumberGenerator.GetBytes(16);
            byte[] packageKey = RandomNumberGenerator.GetBytes(32);
            byte[] verifierInput = RandomNumberGenerator.GetBytes(16);

            PasswordEncryptorBlobs blobs = AgileKeyDerivation.DeriveWriteBlobs(
                PasswordEncryptor(pwSalt, default), Password, packageKey, verifierInput);
            AgileDescriptor descriptor = new(
                KeyData(RandomNumberGenerator.GetBytes(16)), PasswordEncryptor(pwSalt, blobs), [], []);

            ExcelEncryptionException exception = Assert.Throws<ExcelEncryptionException>(
                () => AgileKeyDerivation.DeriveIntermediateKey(descriptor, "wrong"));
            Assert.Equal(ExcelEncryptionReason.PasswordIncorrect, exception.Reason);
        }

        [Fact]
        public void WrapHmac_Output_Is_Readable_By_UnwrapHmac()
        {
            byte[] keySalt = RandomNumberGenerator.GetBytes(16);
            byte[] packageKey = RandomNumberGenerator.GetBytes(32);
            byte[] hmacKey = RandomNumberGenerator.GetBytes(64);
            byte[] hmacValue = RandomNumberGenerator.GetBytes(64);
            CryptoParameters keyData = KeyData(keySalt);

            (byte[] wrappedKey, byte[] wrappedValue) = AgileKeyDerivation.WrapHmac(keyData, packageKey, hmacKey, hmacValue);
            AgileDescriptor descriptor = new(keyData, PasswordEncryptor(keySalt, default), wrappedKey, wrappedValue);

            (byte[] unwrappedKey, byte[] unwrappedValue) = AgileKeyDerivation.UnwrapHmac(descriptor, packageKey);

            Assert.Equal(hmacKey, unwrappedKey);
            Assert.Equal(hmacValue, unwrappedValue);
        }

        [Fact]
        public void EncryptNoPadding_MisalignedPlaintext_Throws()
        {
            byte[] key = RandomNumberGenerator.GetBytes(32);
            byte[] iv = RandomNumberGenerator.GetBytes(16);

            Assert.Throws<ArgumentException>(() => AgileKeyDerivation.EncryptNoPadding(new byte[17], key, iv));
        }
    }
}
