using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;
using ExcelReader.Core.Crypto;

namespace ExcelReader.Tests
{
    // CbcSegmentCipher reuses one ICryptoTransform across every segment and corrects the first
    // block for the IV the segment is actually encrypted under. The oracle for all of it is the
    // one-shot Aes.DecryptCbc/EncryptCbc API it replaced: these tests assert byte-for-byte
    // equality against that, segment by segment, which is the only thing that makes the
    // chaining-state fix-up safe to rely on in a security path.
    public class CbcSegmentCipherTests
    {
        private const int SegmentSize = 4096;
        private const int BlockSize = 16;

        // Three full segments plus a short-but-block-aligned tail, which is the shape a real
        // package ends on.
        private static readonly int[] SegmentLengths = [SegmentSize, SegmentSize, SegmentSize, 32];

        // Deliberately not cryptographically secure — CA5394 doesn't apply: a seeded, reproducible
        // byte pattern standing in for keys, IVs and payloads that are never used to protect
        // anything. A failing case has to be replayable from its seed.
        [SuppressMessage("Security", "CA5394:Do not use insecure randomness",
            Justification = "Seeded test vectors, not randomness protecting anything.")]
        private static byte[] Random(int length, int seed)
        {
            byte[] buffer = new byte[length];
            new Random(seed).NextBytes(buffer);
            return buffer;
        }

        // The per-segment IVs ECMA-376 derives from the salt and the segment index; their only
        // relevant property here is that they differ per segment, so random stands in for them.
        private static byte[][] Ivs(int count, int seed)
        {
            byte[][] ivs = new byte[count][];
            for (int i = 0; i < count; i++)
            {
                ivs[i] = Random(BlockSize, seed + i);
            }
            return ivs;
        }

        private static byte[] OneShotDecrypt(byte[] key, byte[] iv, byte[] cipher)
        {
            using Aes aes = Aes.Create();
            aes.Key = key;
            return aes.DecryptCbc(cipher, iv, PaddingMode.None);
        }

        private static byte[] OneShotEncrypt(byte[] key, byte[] iv, byte[] plain)
        {
            using Aes aes = Aes.Create();
            aes.Key = key;
            return aes.EncryptCbc(plain, iv, PaddingMode.None);
        }

        [Fact]
        public void Should_Match_OneShot_When_Decrypting_Every_Segment()
        {
            byte[] key = Random(32, 1);
            byte[][] ivs = Ivs(SegmentLengths.Length, 100);
            using CbcSegmentCipher cipher = CbcSegmentCipher.CreateDecryptor(key);

            for (int i = 0; i < SegmentLengths.Length; i++)
            {
                byte[] ciphertext = Random(SegmentLengths[i], 200 + i);
                byte[] actual = new byte[ciphertext.Length];
                cipher.Decrypt(ciphertext, ivs[i], actual);
                Assert.Equal(OneShotDecrypt(key, ivs[i], ciphertext), actual);
            }
        }

        // DecryptedPackageStream seeks, so segments arrive in whatever order the ZIP reader asks
        // for them. The fix-up compensates from the state the cipher tracks itself, never from an
        // assumption about which segment came before, so the order must not matter.
        [Fact]
        public void Should_Match_OneShot_When_Segments_Arrive_Out_Of_Order()
        {
            byte[] key = Random(32, 2);
            byte[][] ivs = Ivs(SegmentLengths.Length, 300);
            byte[][] ciphertexts = new byte[SegmentLengths.Length][];
            for (int i = 0; i < SegmentLengths.Length; i++)
            {
                ciphertexts[i] = Random(SegmentLengths[i], 400 + i);
            }

            using CbcSegmentCipher cipher = CbcSegmentCipher.CreateDecryptor(key);
            foreach (int i in (int[])[3, 1, 0, 2, 1, 3, 3, 0])
            {
                byte[] actual = new byte[ciphertexts[i].Length];
                cipher.Decrypt(ciphertexts[i], ivs[i], actual);
                Assert.Equal(OneShotDecrypt(key, ivs[i], ciphertexts[i]), actual);
            }
        }

        [Fact]
        public void Should_Match_OneShot_When_Encrypting_Every_Segment()
        {
            byte[] key = Random(32, 3);
            byte[][] ivs = Ivs(SegmentLengths.Length, 500);
            using CbcSegmentCipher cipher = CbcSegmentCipher.CreateEncryptor(key);

            for (int i = 0; i < SegmentLengths.Length; i++)
            {
                byte[] plaintext = Random(SegmentLengths[i], 600 + i);
                // Encrypt XORs the first block of its input in place, so the oracle needs its own copy.
                byte[] expected = OneShotEncrypt(key, ivs[i], plaintext);
                byte[] actual = new byte[plaintext.Length];
                cipher.Encrypt(plaintext, ivs[i], actual);
                Assert.Equal(expected, actual);
            }
        }

        [Fact]
        public void Should_Round_Trip_When_Encrypted_Then_Decrypted()
        {
            byte[] key = Random(32, 4);
            byte[][] ivs = Ivs(SegmentLengths.Length, 700);
            using CbcSegmentCipher encryptor = CbcSegmentCipher.CreateEncryptor(key);
            using CbcSegmentCipher decryptor = CbcSegmentCipher.CreateDecryptor(key);

            for (int i = 0; i < SegmentLengths.Length; i++)
            {
                byte[] plaintext = Random(SegmentLengths[i], 800 + i);
                byte[] expected = plaintext.AsSpan().ToArray();
                byte[] ciphertext = new byte[plaintext.Length];
                encryptor.Encrypt(plaintext, ivs[i], ciphertext);

                byte[] roundTripped = new byte[ciphertext.Length];
                decryptor.Decrypt(ciphertext, ivs[i], roundTripped);
                Assert.Equal(expected, roundTripped);
            }
        }
    }
}
