using System.Buffers;
using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Crypto
{
    // Decrypts the "EncryptedPackage" CFB stream of an encrypted OOXML workbook on demand, in
    // 4096-byte plaintext segments, so ZipArchive can seek/read the decrypted ZIP without the whole
    // package ever being materialized at once. This is the oracle-tested piece of the pipeline: with
    // no writer in this codebase to round-trip against, byte-exact agreement with msoffcrypto-tool's
    // independently produced plaintext (see EncryptedFixtures/DecryptedPackageStreamTests) is the only
    // correctness signal available.
    //
    // Only agile encryption (AES-CBC per segment) is implemented — EncryptionDescriptor.Parse never
    // yields a concrete descriptor for standard encryption yet (see its own remarks), so there is
    // nothing else to dispatch on here.
    internal sealed class DecryptedPackageStream : Stream
    {
        private const int SegmentSize = 4096;
        private const int PrefixSize = 8;
        // AES's block size is fixed at 128 bits regardless of key length (128/192/256), so this is a
        // true constant, not something read off the descriptor.
        private const int CipherBlockSize = 16;

        // Borrowed only long enough for Create's own reads/derivation; ownership of the ciphertext
        // view below is what actually gets disposed.
        private readonly Stream _view;
        private readonly AgileDescriptor _descriptor;
        private readonly byte[] _key;
        private readonly long _length;
        // Rented once for the stream's lifetime; ZipArchive's access pattern (read sequentially
        // within one entry, occasionally seek) makes a single cached segment enough — see the design
        // doc's remark that a multi-segment cache would be speculative.
        private readonly byte[] _segmentCache;
        // One AES instance for the stream's lifetime instead of one per 4 KB segment. Aes.Create()
        // plus a Key assignment plus CreateDecryptor() ran once per segment - ~25,600 times on a
        // 100 MB package - rebuilding an identical key schedule every round. Only the IV differs
        // between segments, and DecryptCbc takes that per call.
        private readonly Aes _aes;
        // Scratch for the per-segment IV, so deriving it costs no allocation either (see
        // AgileKeyDerivation.SegmentIv's span overload).
        private readonly byte[] _iv;
        // Likewise built once: a per-segment IncrementalHash is a BCryptCreateHash per 4 KB.
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
                // A prefix claiming more plaintext than the container holds ciphertext for, or more
                // than the caller's byte budget allows, is a crafted file - reject it before it sizes
                // a buffer.
                if (declared < 0 || declared > view.Length - PrefixSize)
                {
                    throw new InvalidDataException("The encrypted package's declared size exceeds its ciphertext.");
                }
                // Every segment is decrypted with PaddingMode.None (see DecryptSegment), which requires
                // a whole number of AES blocks. Real files satisfy this by construction (each segment's
                // ciphertext is either a full 4096-byte multiple of 16, or the final segment's ciphertext
                // padded up to the next 16-byte boundary), but the CFB entry size and this prefix are
                // both attacker-controlled - a crafted file can make the total ciphertext land on a
                // non-block boundary, which would otherwise surface as a raw ArgumentOutOfRangeException
                // from TransformBlock deep inside EnsureSegment instead of a graceful rejection here. Same
                // rationale as AgileKeyDerivation.DecryptNoPadding's block-alignment check.
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
                    // Opt-in only here: verifying needs a full pass over the ciphertext before the
                    // first row, unlike the memory path (EncryptedPackageOpener.DecryptToMemory) where
                    // it's nearly free because everything is already decrypted — see
                    // ExcelReaderOptions.VerifyEncryptedIntegrity's remarks. A descriptor with no
                    // dataIntegrity element (AgileDescriptor.HasDataIntegrity) makes opting in a
                    // documented no-op rather than a spurious failure.
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
            // Fast path: the whole requested range already sits inside the cached segment, so this
            // completes synchronously with no state machine allocated (the three-tier convention
            // documented in ARCHITECTURE.md).
            if (_cachedSegment == segment)
            {
                int toCopy = (int)Math.Min(Math.Min(buffer.Length, SegmentSize - offset), _length - _position);
                _segmentCache.AsSpan(offset, toCopy).CopyTo(buffer.Span);
                _position += toCopy;
                return new ValueTask<int>(toCopy);
            }

            return ReadSlowAsync(buffer, cancellationToken);
        }

        // Holds the awaiting refill loop so the fast path above never allocates a state machine.
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

        // Returns immediately when the requested segment is already cached - the fast path that makes
        // ZipArchive's mostly-sequential access pattern cheap. Otherwise seeks the ciphertext view to
        // the segment's offset, reads up to SegmentSize bytes (the final segment is shorter - its
        // ciphertext is only padded out to the next 16-byte block boundary, not a full segment), and
        // decrypts in place into the cache.
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

        // The ciphertext offset/length for segment `index`, bounded by however much ciphertext the
        // view actually holds from that point on - always a multiple of the 16-byte AES block size,
        // since every segment boundary (a multiple of 4096) already is one, and Create already
        // rejected a container whose total ciphertext length (view.Length - PrefixSize) is not itself
        // block-aligned.
        private (long Offset, int Length) SegmentCiphertextRange(int index)
        {
            long segmentStart = (long)index * SegmentSize;
            long cipherOffset = PrefixSize + segmentStart;
            int cipherLen = (int)Math.Min(SegmentSize, _view.Length - cipherOffset);
            return (cipherOffset, cipherLen);
        }

        // AES-CBC, no padding, decrypting straight into the pooled segment cache. DecryptCbc rather
        // than an ICryptoTransform: it takes the IV per call, which is what lets the cipher object and
        // its key schedule be built once in Create instead of rebuilt per segment, and it writes into
        // a caller-owned span, so there is no per-segment output array either. No padding to strip -
        // the final segment's short plaintext is handled by Length bounding the copy in
        // Read/ReadSlowAsync.
        private void DecryptSegment(int index, byte[] cipherBuf, int cipherLen)
        {
            AgileKeyDerivation.SegmentIv(_descriptor, index, _ivHasher, _iv);
            _aes.DecryptCbc(cipherBuf.AsSpan(0, cipherLen), _iv, _segmentCache.AsSpan(0, cipherLen), PaddingMode.None);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "DecryptedPackageStream owns the ciphertext view for the duration of the decrypted read.")]
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
