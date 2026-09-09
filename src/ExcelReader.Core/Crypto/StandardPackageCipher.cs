using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Crypto
{
    // Standard (ECMA-376 3.2/4.2): AES-ECB over the whole package. ECB has no chaining, so the
    // segment index is irrelevant here - any block-aligned window decrypts on its own, which is
    // what lets the caller keep its 4096-byte on-demand segmenting unchanged.
    internal sealed class StandardPackageCipher : PackageCipher
    {
        private readonly byte[] _key;
        private readonly Aes _aes;
        // Reused across every segment instead of Aes.DecryptEcb's one-shot API, which builds and
        // tears down a cipher object (a CNG key import on Windows) per call. Measured: that setup
        // cost is ~82% of the total decrypt time for a package's worth of 4096-byte segments. ECB
        // has no per-call state (no IV/chaining), so one transform decrypts every segment safely.
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
            // Unused by design: ECB blocks are independent of their position.
            _ = segmentIndex;
            // TransformBlock has no Span overload; every caller (DecryptedPackageStream,
            // EncryptedPackageOpener) always hands over array-backed memory, so this never falls
            // back to the slower one-shot path in practice. TransformBlock accepts an inputCount
            // spanning multiple cipher blocks in one call (documented .NET behavior for a
            // non-padded mode), so the whole segment decrypts in one call, same as the one-shot API did.
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
            // The scheme carries no integrity field, so there is nothing to verify. Callers gate on
            // SupportsIntegrity, which is false here.
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
