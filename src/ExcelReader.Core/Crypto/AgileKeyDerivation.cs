using System.Buffers;
using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Crypto
{
    // ECMA-376/[MS-OFFCRYPTO] agile encryption key derivation.
    internal static class AgileKeyDerivation
    {
        // [MS-OFFCRYPTO] 2.3.4.10 password key encryptor block keys.
        private static ReadOnlySpan<byte> BlockVerifierHashInput => [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79];
        private static ReadOnlySpan<byte> BlockVerifierHashValue => [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e];
        private static ReadOnlySpan<byte> BlockKeyValue          => [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6];
        // [MS-OFFCRYPTO] 2.3.4.14 data integrity block keys.
        private static ReadOnlySpan<byte> BlockHmacKey            => [0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6];
        private static ReadOnlySpan<byte> BlockHmacValue          => [0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33];

        // SHA-512, the longest hash the descriptor can name.
        private const int MaxHashLength = 64;

        internal static byte[] DeriveIntermediateKey(AgileDescriptor d, ReadOnlySpan<char> password)
        {
            CryptoParameters p = d.PasswordEncryptor;
            byte[] hFinal = PasswordHash(p, password);
            try
            {
                VerifyPassword(p, hFinal);

                // encryptedKeyValue's IV is the encryptor's own salt, not a hash of it.
                byte[] ivSalt = NormalizeToLength(p.SaltValue, p.BlockSize);
                byte[] raw = DecryptNoPadding(p.EncryptedKeyValue, BlockKey(p, hFinal, BlockKeyValue), ivSalt);

                // encryptedKeyValue is untrusted, pre-authentication input: a producer can wrap fewer
                // bytes than keyBits/8 declares, so reject that explicitly rather than let the slice
                // below throw a raw BCL exception.
                int keyLen = d.KeyData.KeyBits / 8;
                if (raw.Length < keyLen)
                {
                    throw new InvalidDataException(
                        "The encryption descriptor's encryptedKeyValue decrypted to fewer bytes than the declared key size.");
                }
                return raw.Length == keyLen ? raw : raw[..keyLen];
            }
            finally
            {
                CryptographicOperations.ZeroMemory(hFinal);
            }
        }

        internal static byte[] SegmentIv(AgileDescriptor d, int segmentIndex)
        {
            byte[] iv = new byte[d.KeyData.BlockSize];
            using IncrementalHash hasher = CreateHasher(d.KeyData.Hash);
            SegmentIv(d, segmentIndex, hasher, iv);
            return iv;
        }

        // Allocation-free twin for the per-segment decryption loops: the array-returning overload
        // above allocates and builds a fresh IncrementalHash per segment, which a 100 MB package
        // calls 25,600 times.
        internal static void SegmentIv(AgileDescriptor d, int segmentIndex, IncrementalHash hasher, Span<byte> destination)
        {
            CryptoParameters keyData = d.KeyData;
            Span<byte> input = stackalloc byte[MaxHashLength + 4];
            if (keyData.SaltValue.Length + 4 > input.Length)
            {
                // Descriptor is untrusted; fall back rather than overrun the stack buffer.
                NormalizeToLength(SegmentIvSlow(keyData, segmentIndex), keyData.BlockSize).CopyTo(destination);
                return;
            }
            keyData.SaltValue.CopyTo(input);
            BinaryPrimitives.WriteUInt32LittleEndian(input[keyData.SaltValue.Length..], unchecked((uint)segmentIndex));

            Span<byte> hash = stackalloc byte[MaxHashLength];
            int hashLen = HashLength(keyData.Hash);
            hasher.AppendData(input[..(keyData.SaltValue.Length + 4)]);
            hasher.GetHashAndReset(hash[..hashLen]);
            NormalizeInto(hash[..hashLen], destination);
        }

        private static byte[] SegmentIvSlow(CryptoParameters keyData, int segmentIndex)
        {
            Span<byte> counter = stackalloc byte[4];
            BinaryPrimitives.WriteUInt32LittleEndian(counter, unchecked((uint)segmentIndex));
            return HashTwo(keyData.Hash, keyData.SaltValue, counter);
        }

        // Created once by the decryption loop and handed to every SegmentIv call.
        internal static IncrementalHash CreateHasher(HashKind kind)
        {
            return IncrementalHash.CreateHash(AlgorithmName(kind));
        }

        internal static (byte[] Key, byte[] Value) UnwrapHmac(AgileDescriptor d, ReadOnlySpan<byte> intermediateKey)
        {
            CryptoParameters keyData = d.KeyData;
            byte[] ivKey = NormalizeToLength(HashTwo(keyData.Hash, keyData.SaltValue, BlockHmacKey), keyData.BlockSize);
            byte[] ivValue = NormalizeToLength(HashTwo(keyData.Hash, keyData.SaltValue, BlockHmacValue), keyData.BlockSize);

            byte[] hmacKey = DecryptNoPadding(d.EncryptedHmacKey, intermediateKey, ivKey);
            byte[] hmacValue = DecryptNoPadding(d.EncryptedHmacValue, intermediateKey, ivValue);

            // Wrapped plaintexts are the hash's native output length in both cases; the producer may
            // pad the ciphertext out to a blockSize multiple.
            int hashLen = HashLength(keyData.Hash);
            return (
                hmacKey.Length > hashLen ? hmacKey[..hashLen] : hmacKey,
                hmacValue.Length > hashLen ? hmacValue[..hashLen] : hmacValue);
        }

        internal static int HashLength(HashKind kind)
        {
            return kind switch
            {
                HashKind.Sha1 => 20,
                HashKind.Sha256 => 32,
                HashKind.Sha384 => 48,
                HashKind.Sha512 => 64,
                _ => throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    $"Unsupported hash algorithm '{kind}' in the encryption descriptor."),
            };
        }

        private static HashAlgorithmName AlgorithmName(HashKind kind)
        {
            return kind switch
            {
                HashKind.Sha1 => HashAlgorithmName.SHA1,
                HashKind.Sha256 => HashAlgorithmName.SHA256,
                HashKind.Sha384 => HashAlgorithmName.SHA384,
                HashKind.Sha512 => HashAlgorithmName.SHA512,
                _ => throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    $"Unsupported hash algorithm '{kind}' in the encryption descriptor."),
            };
        }

        // H0 = H(salt || UTF16LE(password)), then spinCount iterations of Hn = H(LE32(n) || Hn-1).
        private static byte[] PasswordHash(CryptoParameters p, ReadOnlySpan<char> password)
        {
            HashAlgorithmName name = AlgorithmName(p.Hash);
            int hashLen = HashLength(p.Hash);
            int pwByteLen = checked(password.Length * 2);
            byte[] pwRented = pwByteLen == 0 ? [] : ArrayPool<byte>.Shared.Rent(pwByteLen);

            // Two rotating LE32(n) || Hn-1 buffers, laid out contiguously so each spin iteration
            // (commonly 100,000) costs one AppendData instead of two, alternating to avoid copying
            // the result back into the input.
            Span<byte> bufferA = stackalloc byte[MaxHashLength + 4];
            Span<byte> bufferB = stackalloc byte[MaxHashLength + 4];
            try
            {
                Span<byte> pwBytes = pwRented.AsSpan(0, pwByteLen);
                Encoding.Unicode.GetBytes(password, pwBytes);

                using IncrementalHash hasher = IncrementalHash.CreateHash(name);
                hasher.AppendData(p.SaltValue);
                hasher.AppendData(pwBytes);
                hasher.GetHashAndReset(bufferA[4..(4 + hashLen)]);

                Span<byte> current = bufferA;
                Span<byte> next = bufferB;
                for (int i = 0; i < p.SpinCount; i++)
                {
                    BinaryPrimitives.WriteUInt32LittleEndian(current, unchecked((uint)i));
                    hasher.AppendData(current[..(4 + hashLen)]);
                    hasher.GetHashAndReset(next[4..(4 + hashLen)]);
                    // No tuple swap: Span<byte> is a ref struct and cannot be a tuple element.
                    Span<byte> previous = current;
                    current = next;
                    next = previous;
                }
                return current.Slice(4, hashLen).ToArray();
            }
            finally
            {
                CryptographicOperations.ZeroMemory(pwRented.AsSpan(0, pwByteLen));
                if (pwByteLen != 0)
                {
                    ArrayPool<byte>.Shared.Return(pwRented);
                }
                CryptographicOperations.ZeroMemory(bufferA);
                CryptographicOperations.ZeroMemory(bufferB);
            }
        }

        // Hfinal = H(hFinal || blockKey), sized to p.KeyBits / 8: truncated when longer, padded with
        // 0x36 (not zero) when shorter.
        private static byte[] BlockKey(CryptoParameters p, ReadOnlySpan<byte> hFinal, ReadOnlySpan<byte> blockKey)
        {
            byte[] combined = HashTwo(p.Hash, hFinal, blockKey);
            int keyLen = p.KeyBits / 8;
            return NormalizeToLength(combined, keyLen);
        }

        // Checks that hashing the decrypted verifier input reproduces the decrypted verifier value.
        private static void VerifyPassword(CryptoParameters p, ReadOnlySpan<byte> hFinal)
        {
            byte[] ivSalt = NormalizeToLength(p.SaltValue, p.BlockSize);
            byte[] key1 = BlockKey(p, hFinal, BlockVerifierHashInput);
            byte[] key2 = BlockKey(p, hFinal, BlockVerifierHashValue);

            byte[] verifierInput = DecryptNoPadding(p.EncryptedVerifierHashInput, key1, ivSalt);
            byte[] verifierValue = DecryptNoPadding(p.EncryptedVerifierHashValue, key2, ivSalt);
            byte[] actual = HashOne(p.Hash, verifierInput);

            if (verifierValue.Length < actual.Length
                || !CryptographicOperations.FixedTimeEquals(actual, verifierValue.AsSpan(0, actual.Length)))
            {
                throw new ExcelEncryptionException(ExcelEncryptionReason.PasswordIncorrect,
                    "The supplied password does not match this workbook.");
            }
        }

        private static byte[] DecryptNoPadding(ReadOnlySpan<byte> ciphertext, ReadOnlySpan<byte> key, ReadOnlySpan<byte> iv)
        {
            if (ciphertext.Length == 0)
            {
                return [];
            }
            using Aes aes = Aes.Create();
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            // EncryptionInfo is untrusted input; reject a misaligned ciphertext explicitly rather than
            // let TransformFinalBlock throw a raw CryptographicException.
            int blockBytes = aes.BlockSize / 8;
            if (ciphertext.Length % blockBytes != 0)
            {
                throw new InvalidDataException(
                    "The encryption descriptor contains a ciphertext whose length is not a multiple of the cipher block size.");
            }
            aes.Key = key.ToArray();
            aes.IV = iv.ToArray();
            using ICryptoTransform decryptor = aes.CreateDecryptor();
            return decryptor.TransformFinalBlock(ciphertext.ToArray(), 0, ciphertext.Length);
        }

        private static byte[] HashOne(HashKind kind, ReadOnlySpan<byte> data)
        {
            return HashTwo(kind, data, default);
        }

        private static byte[] HashTwo(HashKind kind, ReadOnlySpan<byte> a, ReadOnlySpan<byte> b)
        {
            byte[] result = new byte[HashLength(kind)];
            using IncrementalHash hasher = IncrementalHash.CreateHash(AlgorithmName(kind));
            hasher.AppendData(a);
            hasher.AppendData(b);
            hasher.GetHashAndReset(result);
            return result;
        }

        // Span twin of NormalizeToLength, for callers that already own the destination.
        private static void NormalizeInto(ReadOnlySpan<byte> value, Span<byte> destination)
        {
            if (value.Length >= destination.Length)
            {
                value[..destination.Length].CopyTo(destination);
                return;
            }
            value.CopyTo(destination);
            destination[value.Length..].Fill(0x36);
        }

        private static byte[] NormalizeToLength(ReadOnlySpan<byte> value, int length)
        {
            if (value.Length == length)
            {
                return value.ToArray();
            }
            byte[] result = new byte[length];
            if (value.Length > length)
            {
                value[..length].CopyTo(result);
            }
            else
            {
                value.CopyTo(result);
                result.AsSpan(value.Length).Fill(0x36);
            }
            return result;
        }
    }
}
