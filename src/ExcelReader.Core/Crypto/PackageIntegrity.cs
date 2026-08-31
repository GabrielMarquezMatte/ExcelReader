using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Crypto
{
    // The HMAC is computed over the whole "EncryptedPackage" stream, including its 8-byte
    // plaintext-size prefix — not just the ciphertext that follows it.
    internal static class PackageIntegrity
    {
        private const int StreamBufferSize = 81920;

        // `ciphertext` must be a view over the WHOLE EncryptedPackage stream (prefix included).
        // Position is saved and restored so callers can invoke this mid-construction.
        internal static void Verify(Stream ciphertext, AgileDescriptor d, ReadOnlySpan<byte> intermediateKey)
        {
            (byte[] key, byte[] expected) = AgileKeyDerivation.UnwrapHmac(d, intermediateKey);
            try
            {
                long start = ciphertext.Position;
                ciphertext.Position = 0;
                byte[] actual = ComputeHmac(ciphertext, d.KeyData.Hash, key);
                ciphertext.Position = start;
                if (!CryptographicOperations.FixedTimeEquals(actual, expected))
                {
                    throw new ExcelEncryptionException(ExcelEncryptionReason.IntegrityFailure,
                        "The encrypted package failed its integrity check: it is corrupt or has been modified.");
                }
            }
            finally
            {
                CryptographicOperations.ZeroMemory(key);
            }
        }

        [SuppressMessage("Security", "CA5350:Do Not Use Weak Cryptographic Algorithms",
            Justification = "The hash algorithm is dictated by the workbook's own EncryptionInfo descriptor, not chosen here; verifying a file written with SHA1 requires HMACSHA1.")]
        [SuppressMessage("Major Code Smell", "S4790:Use a stronger hashing algorithm",
            Justification = "The hash algorithm is dictated by the workbook's own EncryptionInfo descriptor, not chosen here; verifying a file written with SHA1 requires HMACSHA1.")]
        internal static IncrementalHash CreateHmac(HashKind kind, byte[] key)
        {
            HashAlgorithmName name = kind switch
            {
                HashKind.Sha1 => HashAlgorithmName.SHA1,
                HashKind.Sha256 => HashAlgorithmName.SHA256,
                HashKind.Sha384 => HashAlgorithmName.SHA384,
                HashKind.Sha512 => HashAlgorithmName.SHA512,
                _ => throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    $"Unsupported hash algorithm '{kind}' in the encryption descriptor."),
            };
            return IncrementalHash.CreateHMAC(name, key);
        }

        private static byte[] ComputeHmac(Stream source, HashKind kind, byte[] key)
        {
            using IncrementalHash hmac = CreateHmac(kind, key);
            byte[] buffer = ArrayPool<byte>.Shared.Rent(StreamBufferSize);
            try
            {
                int read;
                while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
                {
                    hmac.AppendData(buffer.AsSpan(0, read));
                }
                return hmac.GetHashAndReset();
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }
    }
}
