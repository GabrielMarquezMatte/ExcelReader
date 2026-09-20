using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;
using System.Text;

namespace ExcelReader.Core.Crypto
{
    internal static class StandardKeyDerivation
    {
        private const int IterationCount = 50_000;
        private const int Sha1Length = 20;
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

                BinaryPrimitives.WriteUInt32LittleEndian(counter, 0);
                sha1.AppendData(hash);
                sha1.AppendData(counter);
                sha1.GetHashAndReset(hash);

                return ExpandToKeyLength(sha1, hash, keyBits / 8);
            }
            finally
            {
                CryptographicOperations.ZeroMemory(passwordBytes);
                CryptographicOperations.ZeroMemory(hash);
            }
        }

        internal static bool VerifyPassword(byte[] key, ReadOnlySpan<byte> encryptedVerifier,
            ReadOnlySpan<byte> encryptedVerifierHash)
        {
            if (encryptedVerifier.Length != VerifierLength || encryptedVerifierHash.Length < Sha1Length
                || encryptedVerifierHash.Length > PadLength)
            {
                return false;
            }

            using Aes aes = CreateEcb(key);
            Span<byte> verifier = stackalloc byte[VerifierLength];
            Span<byte> expected = stackalloc byte[Sha1Length];
            Span<byte> actual = stackalloc byte[PadLength];
            try
            {
                aes.DecryptEcb(encryptedVerifier, verifier, PaddingMode.None);

                using IncrementalHash sha1 = CreateSha1();
                sha1.AppendData(verifier);
                sha1.GetHashAndReset(expected);

                int wrapped = encryptedVerifierHash.Length - (encryptedVerifierHash.Length % VerifierLength);
                aes.DecryptEcb(encryptedVerifierHash[..wrapped], actual[..wrapped], PaddingMode.None);

                return CryptographicOperations.FixedTimeEquals(expected, actual[..Sha1Length]);
            }
            finally
            {
                CryptographicOperations.ZeroMemory(verifier);
                CryptographicOperations.ZeroMemory(expected);
                CryptographicOperations.ZeroMemory(actual);
            }
        }

        private static byte[] ExpandToKeyLength(IncrementalHash sha1, ReadOnlySpan<byte> hFinal, int keyLength)
        {
            ArgumentOutOfRangeException.ThrowIfGreaterThan(keyLength, Sha1Length * 2);

            Span<byte> buffer = stackalloc byte[PadLength];
            Span<byte> derived = stackalloc byte[Sha1Length * 2];
            try
            {
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
            finally
            {
                CryptographicOperations.ZeroMemory(buffer);
                CryptographicOperations.ZeroMemory(derived);
            }
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
        private static Aes CreateEcb(byte[] key)
        {
            Aes aes = Aes.Create();
            try
            {
                aes.Mode = CipherMode.ECB;
                aes.Padding = PaddingMode.None;
                aes.Key = key;
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
