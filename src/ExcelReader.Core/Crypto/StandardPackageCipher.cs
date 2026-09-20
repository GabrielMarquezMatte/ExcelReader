using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Crypto
{
    internal sealed class StandardPackageCipher : PackageCipher
    {
        private readonly byte[] _key;
        private readonly Aes _aes;
        private readonly ICryptoTransform _decryptor;
        private bool _disposed;

        [SuppressMessage("Security", "CA5358:Review cipher mode usage with cryptographic experts",
            Justification = "ECMA-376 standard encryption encrypts the package with AES-ECB; reading such a " +
                "file requires ECB.")]
        internal StandardPackageCipher(StandardDescriptor descriptor, ReadOnlySpan<char> password)
        {
            byte[] key = StandardKeyDerivation.DeriveKey(password, descriptor.Salt, descriptor.KeyBits);
            if (!StandardKeyDerivation.VerifyPassword(
                    key, descriptor.EncryptedVerifier, descriptor.EncryptedVerifierHash))
            {
                CryptographicOperations.ZeroMemory(key);
                throw new ExcelEncryptionException(ExcelEncryptionReason.PasswordIncorrect,
                    "The supplied password does not match this workbook.");
            }
            _key = key;
            Aes aes = Aes.Create();
            try
            {
                aes.Mode = CipherMode.ECB;
                aes.Padding = PaddingMode.None;
                aes.Key = _key;
                _aes = aes;
                _decryptor = aes.CreateDecryptor();
            }
            catch
            {
                aes.Dispose();
                CryptographicOperations.ZeroMemory(_key);
                throw;
            }
        }

        internal override bool SupportsIntegrity
        {
            get
            {
                return false;
            }
        }

        internal override void DecryptSegment(int segmentIndex, ReadOnlyMemory<byte> cipher, Memory<byte> plain)
        {
            _ = segmentIndex;
            if (MemoryMarshal.TryGetArray(cipher, out ArraySegment<byte> cipherSeg)
                && MemoryMarshal.TryGetArray((ReadOnlyMemory<byte>)plain, out ArraySegment<byte> plainSeg))
            {
                _decryptor.TransformBlock(cipherSeg.Array!, cipherSeg.Offset, cipherSeg.Count, plainSeg.Array!, plainSeg.Offset);
                return;
            }
            _aes.DecryptEcb(cipher.Span, plain.Span, PaddingMode.None);
        }

        internal override void VerifyIntegrity(Stream ciphertextView)
        {
            _ = ciphertextView;
        }

        protected override void Dispose(bool disposing)
        {
            if (!disposing || _disposed)
            {
                return;
            }
            _disposed = true;
            CryptographicOperations.ZeroMemory(_key);
            _decryptor.Dispose();
            _aes.Dispose();
        }
    }
}
