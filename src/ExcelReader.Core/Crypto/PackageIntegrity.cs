using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Crypto
{
    // [MS-OFFCRYPTO] 2.3.4.14 "DataIntegrity Verification (Agile Encryption)": the HMAC is computed
    // over the raw bytes of the whole "EncryptedPackage" CFB stream, INCLUDING its 8-byte
    // plaintext-size prefix — not just the ciphertext that follows it (2.3.4.4 defines the stream's
    // own content as prefix+ciphertext together, and 2.3.4.14 hashes "the content of the
    // EncryptedPackage stream", not a sub-range of it). Skipping the prefix produces a digest that
    // still fails identically to a wrong key, so this was verified against real fixtures
    // (EncryptedIntegrityTests) rather than trusted from the spec text alone.
    internal static class PackageIntegrity
    {
        // Matches the buffer size FileStream itself defaults to; large enough to keep the syscall
        // count reasonable for a hundreds-of-MB workbook, small enough to stay off the LOH.
        private const int StreamBufferSize = 81920;

        // `ciphertext` must be a view over the WHOLE EncryptedPackage stream (prefix included) — see
        // the file header comment. Its position is saved and restored so callers (DecryptedPackageStream.
        // Create, mid-construction) can call this without disturbing where they were.
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

        // Never a reflected algorithm name — same allowlist-only rule as
        // AgileKeyDerivation/EncryptionDescriptor.ParseHash. HMACSHA1 is not a weak *default* here:
        // it's one of four choices the workbook's own descriptor dictates (whatever hashAlgorithm the
        // file that was encrypted actually used), so CA5350/S4790's "prefer a stronger algorithm"
        // guidance doesn't apply — reading a file that was written with SHA1 requires computing SHA1.
        [SuppressMessage("Security", "CA5350:Do Not Use Weak Cryptographic Algorithms",
            Justification = "The hash algorithm is dictated by the workbook's own EncryptionInfo descriptor, not chosen here; verifying a file written with SHA1 requires HMACSHA1.")]
        [SuppressMessage("Major Code Smell", "S4790:Use a stronger hashing algorithm",
            Justification = "The hash algorithm is dictated by the workbook's own EncryptionInfo descriptor, not chosen here; verifying a file written with SHA1 requires HMACSHA1.")]
        // IncrementalHash rather than HMAC.TransformBlock: TransformBlock is the legacy
        // ICryptoTransform surface, which re-validates its arguments and maintains an output-copy
        // contract this caller never uses (it passes a null output buffer), and it forces a final
        // TransformFinalBlock plus a Hash property allocation. AppendData/GetHashAndReset is the same
        // computation without that. Worth about 8% of the verification pass (measured on a 5.9 MB
        // package: 11.7 ms -> 10.7 ms); the rest is the HMAC itself, which is the work being asked for.
        private static byte[] ComputeHmac(Stream source, HashKind kind, byte[] key)
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

            using IncrementalHash hmac = IncrementalHash.CreateHMAC(name, key);
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
