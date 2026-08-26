using System.Buffers;
using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Crypto
{
    // [MS-OFFCRYPTO] 2.3.4.7 "Encryption Key Generation (Agile Encryption)": H0 = H(salt ||
    // UTF16LE(password)); then spinCount iterations of Hn = H(LE32(n) || Hn-1), where the counter is
    // an unsigned 32-bit little-endian value starting at 0 and ending at spinCount-1. A per-purpose
    // key is then Hfinal = H(Hn || blockKey), truncated to the target key length or padded with 0x36
    // bytes when the hash is shorter than the target (2.3.4.7, last paragraph; also the general IV
    // rule in 2.3.4.12 "Initialization Vector Generation"). Every block-key byte constant and the
    // 0x36 padding rule below were cross-checked against the live spec pages (not re-derived from
    // memory) — see the section references on each constant/helper.
    internal static class AgileKeyDerivation
    {
        // [MS-OFFCRYPTO] 2.3.4.10 "PasswordKeyEncryptor Generation (Agile Encryption)".
        private static ReadOnlySpan<byte> BlockVerifierHashInput => [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79];
        private static ReadOnlySpan<byte> BlockVerifierHashValue => [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e];
        private static ReadOnlySpan<byte> BlockKeyValue          => [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6];
        // [MS-OFFCRYPTO] 2.3.4.14 "DataIntegrity Generation (Agile Encryption)".
        private static ReadOnlySpan<byte> BlockHmacKey            => [0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6];
        private static ReadOnlySpan<byte> BlockHmacValue          => [0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33];

        internal static byte[] DeriveIntermediateKey(AgileDescriptor d, ReadOnlySpan<char> password)
        {
            CryptoParameters p = d.PasswordEncryptor;
            byte[] hFinal = PasswordHash(p, password);
            try
            {
                VerifyPassword(p, hFinal);

                // encryptedKeyValue's IV is the encryptor's own salt (2.3.4.10, step 3 of the
                // encryptedKeyValue generation), not a hash of it: the "no blockKey" branch of the
                // general IV rule in 2.3.4.12.
                byte[] ivSalt = NormalizeToLength(p.SaltValue, p.BlockSize);
                byte[] raw = DecryptNoPadding(p.EncryptedKeyValue, BlockKey(p, hFinal, BlockKeyValue), ivSalt);

                // The plaintext wrapped here is defined by spec to be exactly KeyData.keyBits/8 bytes
                // (2.3.4.10, encryptedKeyValue step 1) regardless of how many blockSize-aligned bytes
                // the ciphertext itself carries.
                int keyLen = d.KeyData.KeyBits / 8;
                return raw.Length == keyLen ? raw : raw[..keyLen];
            }
            finally
            {
                CryptographicOperations.ZeroMemory(hFinal);
            }
        }

        internal static byte[] SegmentIv(AgileDescriptor d, int segmentIndex)
        {
            // [MS-OFFCRYPTO] 2.3.4.13 "Data Encryption (Agile Encryption)": the IV for segment n is
            // H(KeyData.saltValue || LE32(n)), normalized per the general 2.3.4.12 rule.
            CryptoParameters keyData = d.KeyData;
            Span<byte> counter = stackalloc byte[4];
            BinaryPrimitives.WriteUInt32LittleEndian(counter, unchecked((uint)segmentIndex));
            byte[] hash = HashTwo(keyData.Hash, keyData.SaltValue, counter);
            return NormalizeToLength(hash, keyData.BlockSize);
        }

        internal static (byte[] Key, byte[] Value) UnwrapHmac(AgileDescriptor d, ReadOnlySpan<byte> intermediateKey)
        {
            // [MS-OFFCRYPTO] 2.3.4.14: encryptedHmacKey/Value are wrapped under the intermediate
            // (document) key, with IVs H(KeyData.saltValue || blockKey) truncated/padded to blockSize.
            CryptoParameters keyData = d.KeyData;
            byte[] ivKey = NormalizeToLength(HashTwo(keyData.Hash, keyData.SaltValue, BlockHmacKey), keyData.BlockSize);
            byte[] ivValue = NormalizeToLength(HashTwo(keyData.Hash, keyData.SaltValue, BlockHmacValue), keyData.BlockSize);

            byte[] hmacKey = DecryptNoPadding(d.EncryptedHmacKey, intermediateKey, ivKey);
            byte[] hmacValue = DecryptNoPadding(d.EncryptedHmacValue, intermediateKey, ivValue);

            // The wrapped plaintexts are, per spec, exactly saltSize bytes (the random HMAC key,
            // 2.3.4.14 step 2) and exactly the hash's native output length (the HMAC digest, step 5) —
            // both possibly padded out to a blockSize multiple by the producer.
            int saltSize = keyData.SaltSize;
            int hashLen = HashLength(keyData.Hash);
            return (
                hmacKey.Length > saltSize ? hmacKey[..saltSize] : hmacKey,
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
                // Never a reflected algorithm name — see EncryptionDescriptor.ParseHash's own remark.
                _ => throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    $"Unsupported hash algorithm '{kind}' in the encryption descriptor."),
            };
        }

        // H0 = H(salt || UTF16LE(password)), then spinCount iterations of Hn = H(LE32(n) || Hn-1).
        // The rented UTF-16 buffer and both rotating hash buffers are zeroed on exit — the same
        // zero-on-exit shape ExcelPassword's remarks describe, just against a chosen HashKind instead
        // of a fixed algorithm, and reused across all spinCount iterations rather than reallocated
        // per iteration (spinCount is commonly 100,000).
        private static byte[] PasswordHash(CryptoParameters p, ReadOnlySpan<char> password)
        {
            HashAlgorithmName name = AlgorithmName(p.Hash);
            int hashLen = HashLength(p.Hash);
            int pwByteLen = checked(password.Length * 2);
            byte[] pwRented = pwByteLen == 0 ? [] : ArrayPool<byte>.Shared.Rent(pwByteLen);
            byte[] bufA = new byte[hashLen];
            byte[] bufB = new byte[hashLen];
            try
            {
                Span<byte> pwBytes = pwRented.AsSpan(0, pwByteLen);
                Encoding.Unicode.GetBytes(password, pwBytes);

                using IncrementalHash hasher = IncrementalHash.CreateHash(name);
                hasher.AppendData(p.SaltValue);
                hasher.AppendData(pwBytes);
                hasher.GetHashAndReset(bufA);

                Span<byte> counter = stackalloc byte[4];
                byte[] current = bufA;
                byte[] next = bufB;
                for (int i = 0; i < p.SpinCount; i++)
                {
                    BinaryPrimitives.WriteUInt32LittleEndian(counter, unchecked((uint)i));
                    hasher.AppendData(counter);
                    hasher.AppendData(current);
                    hasher.GetHashAndReset(next);
                    (current, next) = (next, current);
                }

                return (byte[])current.Clone();
            }
            finally
            {
                CryptographicOperations.ZeroMemory(pwRented.AsSpan(0, pwByteLen));
                if (pwByteLen != 0)
                {
                    ArrayPool<byte>.Shared.Return(pwRented);
                }
                CryptographicOperations.ZeroMemory(bufA);
                CryptographicOperations.ZeroMemory(bufB);
            }
        }

        // Hfinal = H(hFinal || blockKey), sized to p.KeyBits / 8: truncated when the hash is longer,
        // padded with 0x36 bytes when it is shorter ([MS-OFFCRYPTO] 2.3.4.7, final paragraph — NOT
        // zero-padding; cross-checked against the live spec text, since a wrong pad byte here would
        // only bite on a keyBits/hash combination this pass has no fixture to catch it with).
        private static byte[] BlockKey(CryptoParameters p, ReadOnlySpan<byte> hFinal, ReadOnlySpan<byte> blockKey)
        {
            byte[] combined = HashTwo(p.Hash, hFinal, blockKey);
            int keyLen = p.KeyBits / 8;
            return NormalizeToLength(combined, keyLen);
        }

        // Decrypts encryptedVerifierHashInput/Value under keys derived from the two verifier block
        // keys (2.3.4.10), both with IV = the encryptor's own salt, and checks that hashing the
        // decrypted input reproduces the decrypted verifier value.
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

        // AES-CBC, no padding: every agile decryption operation (verifier, key, HMAC key/value,
        // package segments) uses this same shape.
        private static byte[] DecryptNoPadding(ReadOnlySpan<byte> ciphertext, ReadOnlySpan<byte> key, ReadOnlySpan<byte> iv)
        {
            if (ciphertext.Length == 0)
            {
                return [];
            }
            using Aes aes = Aes.Create();
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
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

        // [MS-OFFCRYPTO] 2.3.4.12 "Initialization Vector Generation (Agile Encryption)": truncate to
        // `length` bytes when longer, pad with 0x36 bytes (not zero) when shorter. Also doubles as the
        // generic key-length normalizer for BlockKey, which follows the identical rule.
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
