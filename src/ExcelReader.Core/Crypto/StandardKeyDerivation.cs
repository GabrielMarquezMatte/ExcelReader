using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;
using System.Text;

namespace ExcelReader.Core.Crypto
{
    // [MS-OFFCRYPTO] 2.3.4.7 ECMA-376 standard encryption key derivation. SHA-1 and the 50,000
    // iteration count are fixed by the scheme rather than stated in the file, which is why neither
    // is a parameter and why no spin-count limit applies on this path: there is no file-supplied
    // value for an attacker to inflate.
    internal static class StandardKeyDerivation
    {
        private const int IterationCount = 50_000;
        private const int Sha1Length = 20;
        // The 0x36/0x5C expansion below hashes a fixed 64-byte block, per the spec.
        private const int PadLength = 64;
        private const int VerifierLength = 16;

        internal static byte[] DeriveKey(ReadOnlySpan<char> password, ReadOnlySpan<byte> salt, int keyBits)
        {
            byte[] passwordBytes = new byte[checked(password.Length * 2)];
            Span<byte> hash = stackalloc byte[Sha1Length];
            Span<byte> counter = stackalloc byte[4];
            try
            {
                Encoding.Unicode.GetBytes(password, passwordBytes);
                using IncrementalHash sha1 = CreateSha1();
                sha1.AppendData(salt);
                sha1.AppendData(passwordBytes);
                sha1.GetHashAndReset(hash);

                for (int i = 0; i < IterationCount; i++)
                {
                    BinaryPrimitives.WriteUInt32LittleEndian(counter, (uint)i);
                    sha1.AppendData(counter);
                    sha1.AppendData(hash);
                    sha1.GetHashAndReset(hash);
                }

                // The trailing block index is always zero: standard encryption derives exactly one
                // key, unlike agile, which derives a separate key per block-key constant.
                BinaryPrimitives.WriteUInt32LittleEndian(counter, 0);
                sha1.AppendData(hash);
                sha1.AppendData(counter);
                sha1.GetHashAndReset(hash);

                return ExpandToKeyLength(sha1, hash, keyBits / 8);
            }
            finally
            {
                CryptographicOperations.ZeroMemory(passwordBytes);
            }
        }

        internal static bool VerifyPassword(ReadOnlySpan<byte> key, ReadOnlySpan<byte> encryptedVerifier,
            ReadOnlySpan<byte> encryptedVerifierHash)
        {
            if (encryptedVerifier.Length != VerifierLength || encryptedVerifierHash.Length < Sha1Length)
            {
                return false;
            }

            using Aes aes = CreateEcb(key);
            Span<byte> verifier = stackalloc byte[VerifierLength];
            aes.DecryptEcb(encryptedVerifier, verifier, PaddingMode.None);

            Span<byte> expected = stackalloc byte[Sha1Length];
            using IncrementalHash sha1 = CreateSha1();
            sha1.AppendData(verifier);
            sha1.GetHashAndReset(expected);

            // The wrapped hash is padded out to a whole cipher block; only its leading 20 bytes are
            // the SHA-1 value.
            Span<byte> actual = stackalloc byte[PadLength];
            int wrapped = encryptedVerifierHash.Length - (encryptedVerifierHash.Length % VerifierLength);
            aes.DecryptEcb(encryptedVerifierHash[..wrapped], actual[..wrapped], PaddingMode.None);

            return CryptographicOperations.FixedTimeEquals(expected, actual[..Sha1Length]);
        }

        // X1 = SHA1(hFinal XOR 0x36, padded with 0x36 to 64 bytes); X2 the same with 0x5C. The key
        // is the leading bytes of X1 || X2, so X2 only matters above a 20-byte key.
        private static byte[] ExpandToKeyLength(IncrementalHash sha1, ReadOnlySpan<byte> hFinal, int keyLength)
        {
            Span<byte> buffer = stackalloc byte[PadLength];
            Span<byte> derived = stackalloc byte[Sha1Length * 2];

            buffer.Fill(0x36);
            XorInto(hFinal, buffer);
            sha1.AppendData(buffer);
            sha1.GetHashAndReset(derived[..Sha1Length]);

            buffer.Fill(0x5C);
            XorInto(hFinal, buffer);
            sha1.AppendData(buffer);
            sha1.GetHashAndReset(derived[Sha1Length..]);

            return derived[..keyLength].ToArray();
        }

        private static void XorInto(ReadOnlySpan<byte> value, Span<byte> destination)
        {
            for (int i = 0; i < value.Length; i++)
            {
                destination[i] ^= value[i];
            }
        }

        private static IncrementalHash CreateSha1()
        {
            return IncrementalHash.CreateHash(HashAlgorithmName.SHA1);
        }

        [SuppressMessage("Security", "CA5358:Review cipher mode usage with cryptographic experts",
            Justification = "ECMA-376 standard encryption encrypts the package with AES-ECB; reading such a file requires ECB. No data is encrypted here.")]
        private static Aes CreateEcb(ReadOnlySpan<byte> key)
        {
            Aes aes = Aes.Create();
            try
            {
                aes.Mode = CipherMode.ECB;
                aes.Padding = PaddingMode.None;
                aes.Key = key.ToArray();
                return aes;
            }
            catch
            {
                aes.Dispose();
                throw;
            }
        }
    }
}
