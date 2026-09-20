using System.Security.Cryptography;

namespace ExcelReader.Core.Crypto
{
    internal sealed class AgilePackageCipher : PackageCipher
    {
        private readonly AgileDescriptor _descriptor;
        private readonly byte[] _key;
        private readonly CbcSegmentCipher _cipher;
        private readonly IncrementalHash _ivHasher;
        private readonly byte[] _iv;
        private bool _disposed;

        internal AgilePackageCipher(AgileDescriptor descriptor, ReadOnlySpan<char> password)
        {
            _descriptor = descriptor;
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
