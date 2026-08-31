using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;

namespace ExcelReader.Core.Crypto
{
    /// <summary>The three wrapped values a password key encryptor carries in the descriptor.</summary>
    internal readonly record struct PasswordEncryptorBlobs(
        byte[] EncryptedVerifierHashInput,
        byte[] EncryptedVerifierHashValue,
        byte[] EncryptedKeyValue);

    // Write direction of the agile derivation in AgileKeyDerivation.cs — same block keys, same
    // Norm/BlockKey/PasswordHash helpers, run forwards. Randomness is a parameter, never generated
    // here, so every function below is deterministic and unit-testable.
    internal static partial class AgileKeyDerivation
    {
        internal static PasswordEncryptorBlobs DeriveWriteBlobs(CryptoParameters passwordEncryptor,
            ReadOnlySpan<char> password, ReadOnlySpan<byte> packageKey, ReadOnlySpan<byte> verifierInput)
        {
            byte[] hFinal = PasswordHash(passwordEncryptor, password);
            byte[] keyVerifierInput = BlockKey(passwordEncryptor, hFinal, BlockVerifierHashInput);
            byte[] keyVerifierValue = BlockKey(passwordEncryptor, hFinal, BlockVerifierHashValue);
            byte[] keyPackageKey = BlockKey(passwordEncryptor, hFinal, BlockKeyValue);
            try
            {
                // The verifier IV is the encryptor's own salt, not a hash of it — mirrors VerifyPassword.
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

        internal static (byte[] EncryptedKey, byte[] EncryptedValue) WrapHmac(CryptoParameters keyData,
            ReadOnlySpan<byte> packageKey, ReadOnlySpan<byte> hmacKey, ReadOnlySpan<byte> hmacValue)
        {
            byte[] ivKey = NormalizeToLength(HashTwo(keyData.Hash, keyData.SaltValue, BlockHmacKey), keyData.BlockSize);
            byte[] ivValue = NormalizeToLength(HashTwo(keyData.Hash, keyData.SaltValue, BlockHmacValue), keyData.BlockSize);
            return (EncryptNoPadding(hmacKey, packageKey, ivKey), EncryptNoPadding(hmacValue, packageKey, ivValue));
        }

        // Every value this writer wraps is already a multiple of the 16-byte block under the fixed
        // parameters (verifier input 16, verifier hash 64, package key 32, HMAC key/value 64), so a
        // misaligned plaintext means a parameter changed and the caller's assumptions are stale.
        [SuppressMessage("Security", "CA5401", Justification = "Agile encryption uses derived IVs per ECMA-376.")]
        internal static byte[] EncryptNoPadding(ReadOnlySpan<byte> plaintext, ReadOnlySpan<byte> key, ReadOnlySpan<byte> iv)
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
            aes.Key = key.ToArray();
            aes.IV = iv.ToArray();
            using ICryptoTransform encryptor = aes.CreateEncryptor();
            return encryptor.TransformFinalBlock(plaintext.ToArray(), 0, plaintext.Length);
        }
    }
}
