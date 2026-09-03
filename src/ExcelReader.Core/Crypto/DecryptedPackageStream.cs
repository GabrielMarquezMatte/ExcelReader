using System.Buffers;
using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Crypto
{
    // Decrypts the "EncryptedPackage" CFB stream of an encrypted OOXML workbook on demand, in
    // 4096-byte plaintext segments, so ZipArchive can seek/read the decrypted ZIP without the whole
    // package ever being materialized at once.
    //
    // Only agile encryption (AES-CBC per segment) is implemented; EncryptionDescriptor.Parse never
    // yields a descriptor for standard encryption.
    internal sealed class DecryptedPackageStream : Stream
    {
        private const int SegmentSize = 4096;
        private const int PrefixSize = 8;
        private const int CipherBlockSize = 16;

        // Borrowed only long enough for Create's own reads/derivation; the ciphertext view below is
        // what actually gets disposed.
        private readonly Stream _view;
        private readonly AgileDescriptor _descriptor;
        private readonly byte[] _key;
        private readonly long _length;
        // A single cached segment is enough for ZipArchive's mostly-sequential access pattern.
        private readonly byte[] _segmentCache;
        // Built once for the stream's lifetime, not per 4 KB segment.
        private readonly Aes _aes;
        private readonly byte[] _iv;
        private readonly IncrementalHash _ivHasher;
        private int _cachedSegment = -1;
        private long _position;
        private bool _disposed;

        private DecryptedPackageStream(Stream view, AgileDescriptor descriptor, byte[] key, long length, byte[] segmentCache, Aes aes)
        {
            _view = view;
            _descriptor = descriptor;
            _key = key;
            _length = length;
            _segmentCache = segmentCache;
            _aes = aes;
            _iv = new byte[descriptor.KeyData.BlockSize];
            _ivHasher = AgileKeyDerivation.CreateHasher(descriptor.KeyData.Hash);
        }

        internal static DecryptedPackageStream Create(CfbContainer cfb, EncryptionDescriptor descriptor, ExcelReaderOptions options)
        {
            if (descriptor is not AgileDescriptor agile)
            {
                throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    "Only agile (ECMA-376 4.4) encryption is supported for streaming decryption.");
            }
            if (options.Password is null)
            {
                throw new ExcelEncryptionException(ExcelEncryptionReason.PasswordRequired,
                    "This workbook is encrypted and requires a password.");
            }

            Stream view = cfb.OpenStreamView("EncryptedPackage");
            try
            {
                if (view.Length < PrefixSize)
                {
                    throw new InvalidDataException("The EncryptedPackage stream is truncated.");
                }
                Span<byte> prefix = stackalloc byte[PrefixSize];
                view.ReadExactly(prefix);
                long declared = BinaryPrimitives.ReadInt64LittleEndian(prefix);
                // Reject a crafted file claiming more plaintext than the ciphertext could hold, before
                // it sizes a buffer.
                if (declared < 0 || declared > view.Length - PrefixSize)
                {
                    throw new InvalidDataException("The encrypted package's declared size exceeds its ciphertext.");
                }
                // Segments decrypt with PaddingMode.None, which requires a whole number of AES blocks
                // reject misalignment here rather than let DecryptCbc throw.
                long cipherTotal = view.Length - PrefixSize;
                if (cipherTotal % CipherBlockSize != 0)
                {
                    throw new InvalidDataException(
                        "The encrypted package's ciphertext length is not a multiple of the cipher block size.");
                }
                LimitChecks.ThrowIfEntryLengthExceeds(
                    declared, options.MaxTotalDecompressedBytes, nameof(ExcelReaderOptions.MaxTotalDecompressedBytes));

                byte[] key = AgileKeyDerivation.DeriveIntermediateKey(agile, options.Password.Chars);
                try
                {
                    // Opt-in only: verifying needs a full extra pass over the ciphertext before the
                    // first row, unlike the memory path where everything is already decrypted.
                    if (agile.HasDataIntegrity && options.VerifyEncryptedIntegrity)
                    {
                        PackageIntegrity.Verify(view, agile, key);
                    }
                }
                catch
                {
                    CryptographicOperations.ZeroMemory(key);
                    throw;
                }

                byte[] segmentCache = ArrayPool<byte>.Shared.Rent(SegmentSize);
                Aes aes = Aes.Create();
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.None;
                aes.Key = key;
                return new DecryptedPackageStream(view, agile, key, declared, segmentCache, aes);
            }
            catch
            {
                view.Dispose();
                throw;
            }
        }

        public override bool CanRead => true;

        public override bool CanSeek => true;

        public override bool CanWrite => false;

        public override long Length => _length;

        public override long Position
        {
            get => _position;
            set
            {
                ArgumentOutOfRangeException.ThrowIfNegative(value);
                _position = value;
            }
        }

        public override void Flush()
        {
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return Read(buffer.AsSpan(offset, count));
        }

        public override int Read(Span<byte> buffer)
        {
            int total = 0;
            while (total < buffer.Length && _position < _length)
            {
                int segment = (int)(_position / SegmentSize);
                int offset = (int)(_position % SegmentSize);
                EnsureSegment(segment);
                int toCopy = (int)Math.Min(Math.Min(buffer.Length - total, SegmentSize - offset), _length - _position);
                _segmentCache.AsSpan(offset, toCopy).CopyTo(buffer.Slice(total, toCopy));
                total += toCopy;
                _position += toCopy;
            }
            return total;
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        }

        public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            if (buffer.Length == 0 || _position >= _length)
            {
                return new ValueTask<int>(0);
            }

            int segment = (int)(_position / SegmentSize);
            int offset = (int)(_position % SegmentSize);
            // Fast path: already cached, so this completes synchronously with no state machine.
            if (_cachedSegment == segment)
            {
                int toCopy = (int)Math.Min(Math.Min(buffer.Length, SegmentSize - offset), _length - _position);
                _segmentCache.AsSpan(offset, toCopy).CopyTo(buffer.Span);
                _position += toCopy;
                return new ValueTask<int>(toCopy);
            }

            return ReadSlowAsync(buffer, cancellationToken);
        }

        private async ValueTask<int> ReadSlowAsync(Memory<byte> buffer, CancellationToken cancellationToken)
        {
            int total = 0;
            while (total < buffer.Length && _position < _length)
            {
                int segment = (int)(_position / SegmentSize);
                int offset = (int)(_position % SegmentSize);
                await EnsureSegmentAsync(segment, cancellationToken).ConfigureAwait(false);
                int toCopy = (int)Math.Min(Math.Min(buffer.Length - total, SegmentSize - offset), _length - _position);
                _segmentCache.AsSpan(offset, toCopy).CopyTo(buffer.Span.Slice(total, toCopy));
                total += toCopy;
                _position += toCopy;
            }
            return total;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            long newPosition = origin switch
            {
                SeekOrigin.Begin => offset,
                SeekOrigin.Current => _position + offset,
                SeekOrigin.End => _length + offset,
                _ => throw new ArgumentOutOfRangeException(nameof(origin)),
            };
            if (newPosition < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(offset), offset,
                    "An attempt was made to move the position before the beginning of the stream.");
            }
            _position = newPosition;
            return _position;
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }

        private void EnsureSegment(int index)
        {
            if (_cachedSegment == index)
            {
                return;
            }

            (long cipherOffset, int cipherLen) = SegmentCiphertextRange(index);
            _view.Seek(cipherOffset, SeekOrigin.Begin);
            byte[] cipherBuf = ArrayPool<byte>.Shared.Rent(cipherLen);
            try
            {
                _view.ReadExactly(cipherBuf.AsSpan(0, cipherLen));
                DecryptSegment(index, cipherBuf, cipherLen);
                _cachedSegment = index;
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(cipherBuf);
            }
        }

        private async ValueTask EnsureSegmentAsync(int index, CancellationToken cancellationToken)
        {
            if (_cachedSegment == index)
            {
                return;
            }

            (long cipherOffset, int cipherLen) = SegmentCiphertextRange(index);
            _view.Seek(cipherOffset, SeekOrigin.Begin);
            byte[] cipherBuf = ArrayPool<byte>.Shared.Rent(cipherLen);
            try
            {
                await _view.ReadExactlyAsync(cipherBuf.AsMemory(0, cipherLen), cancellationToken).ConfigureAwait(false);
                DecryptSegment(index, cipherBuf, cipherLen);
                _cachedSegment = index;
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(cipherBuf);
            }
        }

        private (long Offset, int Length) SegmentCiphertextRange(int index)
        {
            long segmentStart = (long)index * SegmentSize;
            long cipherOffset = PrefixSize + segmentStart;
            int cipherLen = (int)Math.Min(SegmentSize, _view.Length - cipherOffset);
            return (cipherOffset, cipherLen);
        }

        private void DecryptSegment(int index, byte[] cipherBuf, int cipherLen)
        {
            AgileKeyDerivation.SegmentIv(_descriptor, index, _ivHasher, _iv);
            _aes.DecryptCbc(cipherBuf.AsSpan(0, cipherLen), _iv, _segmentCache.AsSpan(0, cipherLen), PaddingMode.None);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && !_disposed)
            {
                _disposed = true;
                CryptographicOperations.ZeroMemory(_key);
                CryptographicOperations.ZeroMemory(_segmentCache);
                ArrayPool<byte>.Shared.Return(_segmentCache);
                _aes.Dispose();
                _ivHasher.Dispose();
                _view.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
