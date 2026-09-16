using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;

namespace ExcelReader.Core.Crypto
{
    /// <summary>The three wrapped values a password key encryptor carries in the descriptor.</summary>
    internal readonly record struct PasswordEncryptorBlobs(
        byte[] EncryptedVerifierHashInput,
        byte[] EncryptedVerifierHashValue,
        byte[] EncryptedKeyValue);

    internal static partial class AgileKeyDerivation
    {
        internal static PasswordEncryptorBlobs DeriveWriteBlobs(CryptoParameters passwordEncryptor,
            ReadOnlySpan<char> password, byte[] packageKey, byte[] verifierInput)
        {
            byte[] hFinal = PasswordHash(passwordEncryptor, password);
            byte[] keyVerifierInput = BlockKey(passwordEncryptor, hFinal, BlockVerifierHashInput);
            byte[] keyVerifierValue = BlockKey(passwordEncryptor, hFinal, BlockVerifierHashValue);
            byte[] keyPackageKey = BlockKey(passwordEncryptor, hFinal, BlockKeyValue);
            try
            {
                byte[] ivSalt = NormalizeToLength(passwordEncryptor.SaltValue, passwordEncryptor.BlockSize);
                byte[] verifierHash = HashOne(passwordEncryptor.Hash, verifierInput);
                return new PasswordEncryptorBlobs(
                    EncryptNoPadding(verifierInput, keyVerifierInput, ivSalt),
                    EncryptNoPadding(verifierHash, keyVerifierValue, ivSalt),
                    EncryptNoPadding(packageKey, keyPackageKey, ivSalt));
            }
            finally
            {
                CryptographicOperations.ZeroMemory(hFinal);
                CryptographicOperations.ZeroMemory(keyVerifierInput);
                CryptographicOperations.ZeroMemory(keyVerifierValue);
                CryptographicOperations.ZeroMemory(keyPackageKey);
            }
        }

        internal static (byte[] EncryptedKey, byte[] EncryptedValue) WrapHmac(CryptoParameters keyData, byte[] packageKey, byte[] hmacKey, byte[] hmacValue)
        {
            byte[] ivKey = NormalizeToLength(HashTwo(keyData.Hash, keyData.SaltValue, BlockHmacKey), keyData.BlockSize);
            byte[] ivValue = NormalizeToLength(HashTwo(keyData.Hash, keyData.SaltValue, BlockHmacValue), keyData.BlockSize);
            return (EncryptNoPadding(hmacKey, packageKey, ivKey), EncryptNoPadding(hmacValue, packageKey, ivValue));
        }

        [SuppressMessage("Security", "CA5401", Justification = "Agile encryption uses derived IVs per ECMA-376.")]
        internal static byte[] EncryptNoPadding(byte[] plaintext, byte[] key, byte[] iv)
        {
            using Aes aes = Aes.Create();
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            int blockBytes = aes.BlockSize / 8;
            if (plaintext.Length == 0 || plaintext.Length % blockBytes != 0)
            {
                throw new ArgumentException(
                    $"Plaintext length {plaintext.Length} is not a positive multiple of the {blockBytes}-byte cipher block.",
                    nameof(plaintext));
            }
            aes.Key = key;
            aes.IV = iv;
            using ICryptoTransform encryptor = aes.CreateEncryptor();
            return encryptor.TransformFinalBlock(plaintext, 0, plaintext.Length);
        }
    }
}
