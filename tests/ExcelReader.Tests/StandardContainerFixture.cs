using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;
using ExcelReader.Core.Crypto;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Tests
{
    // Forges a standard-encrypted container from the agile fixture's decrypted plaintext. This is a
    // wiring fixture, not a conformance oracle: it is encrypted by the same derivation that decrypts
    // it. See StandardEncryptionOpenTests' class comment.
    internal static class StandardContainerFixture
    {
        internal const string Password = "hunter2";
        private const int AlgIdAes256 = 0x00006610;

        private static readonly byte[] Salt =
        [
            0x01, 0x08, 0x0F, 0x16, 0x1D, 0x24, 0x2B, 0x32,
            0x39, 0x40, 0x47, 0x4E, 0x55, 0x5C, 0x63, 0x6A,
        ];

        internal static byte[] Forge()
        {
            return Forge(PlainPackage(), Password);
        }

        [SuppressMessage("Security", "CA5358:Review cipher mode usage with cryptographic experts",
            Justification = "ECMA-376 standard encryption encrypts the package with AES-ECB; forging a " +
                "conformant container for this wiring test requires ECB.")]
        [SuppressMessage("Security", "CA5350:Do Not Use Weak Cryptographic Algorithms",
            Justification = "ECMA-376 standard encryption's password verifier hash is SHA-1 by " +
                "definition; forging a conformant container for this wiring test requires SHA-1.")]
        internal static byte[] Forge(byte[] plainPackage, string password)
        {
            byte[] key = StandardKeyDerivation.DeriveKey(password, Salt, keyBits: 256);
            using Aes aes = Aes.Create();
            aes.Mode = CipherMode.ECB;
            aes.Padding = PaddingMode.None;
            aes.Key = key;

            byte[] verifier = new byte[16];
            RandomNumberGenerator.Fill(verifier);
            byte[] verifierHash = new byte[32];
            SHA1.HashData(verifier).CopyTo(verifierHash.AsSpan());

            byte[] encryptedVerifier = aes.EncryptEcb(verifier, PaddingMode.None);
            byte[] encryptedVerifierHash = aes.EncryptEcb(verifierHash, PaddingMode.None);
            byte[] info = BuildInfo(encryptedVerifier, encryptedVerifierHash);

            int padded = (plainPackage.Length + 15) / 16 * 16;
            byte[] block = new byte[padded];
            plainPackage.CopyTo(block.AsSpan());
            byte[] cipherText = aes.EncryptEcb(block, PaddingMode.None);

            byte[] package = new byte[8 + cipherText.Length];
            BinaryPrimitives.WriteInt64LittleEndian(package.AsSpan(0, 8), plainPackage.Length);
            cipherText.CopyTo(package.AsSpan(8));

            using var container = new MemoryStream();
            OleCompoundWriter.Write(container,
            [
                new CfbStreamSpec("EncryptionInfo", info.Length,
                    stream => stream.Write(info),
                    (stream, ct) => stream.WriteAsync(info, ct)),
                new CfbStreamSpec("EncryptedPackage", package.Length,
                    stream => stream.Write(package),
                    (stream, ct) => stream.WriteAsync(package, ct)),
            ]);
            return container.ToArray();
        }

        private static byte[] BuildInfo(byte[] encryptedVerifier, byte[] encryptedVerifierHash)
        {
            byte[] info = StandardInfoBuilder.Build(4, 2, algId: AlgIdAes256, keySize: 256);
            // StandardInfoBuilder writes zeroed salt/verifier placeholders; overwrite them with the
            // real values this container is encrypted under. They are the last 68 bytes: salt(16),
            // encryptedVerifier(16), verifierHashSize(4), encryptedVerifierHash(32).
            int saltOffset = info.Length - 68;
            Salt.CopyTo(info.AsSpan(saltOffset));
            encryptedVerifier.CopyTo(info.AsSpan(saltOffset + 16));
            encryptedVerifierHash.CopyTo(info.AsSpan(saltOffset + 36));
            return info;
        }

        // Reuses the agile fixture's decrypted plaintext, which is a real XLSX package and is
        // already asserted to be a ZIP and larger than one segment.
        private static byte[] PlainPackage()
        {
            return EncryptedFixtures.PlainBytes("agile-aes256-sha512.xlsx");
        }
    }
}
