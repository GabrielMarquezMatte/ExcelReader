using System.Security.Cryptography;

namespace ExcelReader.Core.Crypto
{
    // Agile (ECMA-376 4.4): AES-CBC with a per-segment IV derived from the keyData salt and the
    // segment index, plus the optional dataIntegrity HMAC.
    internal sealed class AgilePackageCipher : PackageCipher
    {
        private readonly AgileDescriptor _descriptor;
        private readonly byte[] _key;
        private readonly CbcSegmentCipher _cipher;
        private readonly IncrementalHash _ivHasher;
        // Rebuilt per segment, but allocated once: a 100 MB package decrypts 25,600 segments.
        private readonly byte[] _iv;
        private bool _disposed;

        internal AgilePackageCipher(AgileDescriptor descriptor, ReadOnlySpan<char> password)
        {
            _descriptor = descriptor;
            // Derives and verifies in one step, so a wrong password throws before anything
            // cipher-shaped is allocated.
            _key = AgileKeyDerivation.DeriveIntermediateKey(descriptor, password);
            CbcSegmentCipher cipher = CbcSegmentCipher.CreateDecryptor(_key);
            try
            {
                _ivHasher = AgileKeyDerivation.CreateHasher(descriptor.KeyData.Hash);
                _cipher = cipher;
            }
            catch
            {
                cipher.Dispose();
                CryptographicOperations.ZeroMemory(_key);
                throw;
            }
            _iv = new byte[descriptor.KeyData.BlockSize];
        }

        internal override bool SupportsIntegrity
        {
            get
            {
                return _descriptor.HasDataIntegrity;
            }
        }

        internal override void DecryptSegment(int segmentIndex, ReadOnlyMemory<byte> cipher, Memory<byte> plain)
        {
            AgileKeyDerivation.SegmentIv(_descriptor, segmentIndex, _ivHasher, _iv);
            _cipher.Decrypt(cipher, _iv, plain);
        }

        internal override void VerifyIntegrity(Stream ciphertextView)
        {
            PackageIntegrity.Verify(ciphertextView, _descriptor, _key);
        }

        protected override void Dispose(bool disposing)
        {
            if (!disposing || _disposed)
            {
                return;
            }
            _disposed = true;
            CryptographicOperations.ZeroMemory(_key);
            _cipher.Dispose();
            _ivHasher.Dispose();
        }
    }
}
