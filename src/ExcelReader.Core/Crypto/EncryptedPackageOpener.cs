using System.Buffers;
using System.Buffers.Binary;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Crypto
{
    // An encrypted OOXML file is a CFB container holding EncryptionInfo + EncryptedPackage, not a
    // ZIP. This turns one into a seekable plaintext-ZIP stream that XlsxReader/XlsbReader consume
    // unchanged.
    internal static class EncryptedPackageOpener
    {
        private const long MaxEncryptionInfoBytes = 64 * 1024;
        private const int SegmentSize = 4096;
        private const int PrefixSize = 8;
        private const int CipherBlockSize = 16;

        // Probes the directory then rewinds; ExcelReaderOptions.Default is used deliberately, since
        // this is a structural yes/no question, not a decrypt.
        internal static bool IsEncryptedContainer(Stream seekableSource)
        {
            long start = seekableSource.Position;
            try
            {
                using CfbContainer cfb = CfbContainer.Parse(seekableSource, ownsSource: false, ExcelReaderOptions.Default);
                return cfb.ContainsStream("EncryptedPackage");
            }
            catch (InvalidDataException)
            {
                return false;
            }
            finally
            {
                seekableSource.Position = start;
            }
        }

        // In-memory twin of IsEncryptedContainer, for Excel.From/Excel.Open's ReadOnlyMemory<byte>
        // overloads.
        internal static bool IsEncryptedMemory(ReadOnlyMemory<byte> container, ExcelReaderOptions options)
        {
            _ = options;
            using Stream source = WrapMemory(container);
            try
            {
                using CfbContainer cfb = CfbContainer.Parse(source, ownsSource: false, ExcelReaderOptions.Default);
                return cfb.ContainsStream("EncryptedPackage");
            }
            catch (InvalidDataException)
            {
                return false;
            }
        }

        // Eager twin of Decrypt for the ReadOnlyMemory<byte> overloads, which never suspend and so
        // cannot use a lazily decrypt-on-demand DecryptedPackageStream.
        internal static ReadOnlyMemory<byte> DecryptToMemory(ReadOnlyMemory<byte> container, ExcelReaderOptions options)
        {
            using Stream source = WrapMemory(container);
            using CfbContainer cfb = CfbContainer.Parse(source, ownsSource: false, options);
            if (options.Password is null)
            {
                throw new ExcelEncryptionException(ExcelEncryptionReason.PasswordRequired,
                    "This workbook is encrypted. Supply ExcelReaderOptions.Password to open it.");
            }
            byte[] info = cfb.ReadStream("EncryptionInfo", MaxEncryptionInfoBytes);
            EncryptionDescriptor descriptor = EncryptionDescriptor.Parse(info, options);
            if (descriptor is not AgileDescriptor agile)
            {
                throw new ExcelEncryptionException(ExcelEncryptionReason.UnsupportedScheme,
                    "Only agile (ECMA-376 4.4) encryption is supported for the in-memory open path.");
            }

            byte[] package = cfb.ReadStream("EncryptedPackage", options.MaxTotalDecompressedBytes);
            if (package.Length < PrefixSize)
            {
                throw new InvalidDataException("The EncryptedPackage stream is truncated.");
            }
            long declared = BinaryPrimitives.ReadInt64LittleEndian(package.AsSpan(0, PrefixSize));
            if (declared < 0 || declared > package.Length - PrefixSize)
            {
                throw new InvalidDataException("The encrypted package's declared size exceeds its ciphertext.");
            }
            long cipherTotal = package.Length - PrefixSize;
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
                if (agile.HasDataIntegrity)
                {
                    using var packageView = new MemoryStream(package, writable: false);
                    PackageIntegrity.Verify(packageView, agile, key);
                }
                return DecryptWholePackage(package, agile, key, declared);
            }
            finally
            {
                CryptographicOperations.ZeroMemory(key);
            }
        }

        private static byte[] DecryptWholePackage(byte[] package, AgileDescriptor agile, byte[] key, long declaredLength)
        {
            int cipherLen = package.Length - PrefixSize;
            byte[] result = new byte[declaredLength];
            byte[] segmentBuffer = ArrayPool<byte>.Shared.Rent(SegmentSize);
            using Aes aes = Aes.Create();
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.None;
            aes.Key = key;
            using IncrementalHash ivHasher = AgileKeyDerivation.CreateHasher(agile.KeyData.Hash);
            Span<byte> iv = stackalloc byte[agile.KeyData.BlockSize];
            try
            {
                int segmentCount = (cipherLen + SegmentSize - 1) / SegmentSize;
                for (int i = 0; i < segmentCount; i++)
                {
                    int segOffset = i * SegmentSize;
                    int segCipherLen = Math.Min(SegmentSize, cipherLen - segOffset);
                    AgileKeyDerivation.SegmentIv(agile, i, ivHasher, iv);
                    aes.DecryptCbc(
                        package.AsSpan(PrefixSize + segOffset, segCipherLen), iv,
                        segmentBuffer.AsSpan(0, segCipherLen), PaddingMode.None);

                    int copyLen = (int)Math.Min(segCipherLen, declaredLength - segOffset);
                    if (copyLen > 0)
                    {
                        Array.Copy(segmentBuffer, 0, result, segOffset, copyLen);
                    }
                }
            }
            finally
            {
                CryptographicOperations.ZeroMemory(segmentBuffer.AsSpan(0, SegmentSize));
                ArrayPool<byte>.Shared.Return(segmentBuffer);
            }
            return result;
        }

        // Avoids a full-buffer copy when `container` is already array-backed.
        private static MemoryStream WrapMemory(ReadOnlyMemory<byte> container)
        {
            return MemoryMarshal.TryGetArray(container, out ArraySegment<byte> segment)
                ? new MemoryStream(segment.Array!, segment.Offset, segment.Count, writable: false)
                : new MemoryStream(container.ToArray(), writable: false);
        }

        // Turns an encrypted OOXML CFB container into a seekable, read-only plaintext-ZIP stream.
        internal static Stream Decrypt(Stream source, bool leaveOpen, ExcelReaderOptions options)
        {
            CfbContainer cfb = CfbContainer.Parse(source, ownsSource: !leaveOpen, options);
            try
            {
                if (options.Password is null)
                {
                    throw new ExcelEncryptionException(ExcelEncryptionReason.PasswordRequired,
                        "This workbook is encrypted. Supply ExcelReaderOptions.Password to open it.");
                }
                byte[] info = cfb.ReadStream("EncryptionInfo", MaxEncryptionInfoBytes);
                EncryptionDescriptor descriptor = EncryptionDescriptor.Parse(info, options);
                DecryptedPackageStream package = DecryptedPackageStream.Create(cfb, descriptor, options);
                // package never stores `cfb`, but its lazy per-segment reads keep seeking back into
                // it — so `cfb` must outlive it, and both need to dispose together.
                return new OwnedDecryptedStream(cfb, package);
            }
            catch
            {
                cfb.Dispose();
                throw;
            }
        }

        // Bundles the decrypted plaintext-ZIP stream with the CFB container it was carved out of, so
        // the pair disposes together: the underlying CFB source stream must stay open for as long as
        // DecryptedPackageStream is still lazily decrypting segments from it.
        private sealed class OwnedDecryptedStream : Stream
        {
            private readonly CfbContainer _cfb;
            private readonly DecryptedPackageStream _inner;
            private bool _disposed;

            internal OwnedDecryptedStream(CfbContainer cfb, DecryptedPackageStream inner)
            {
                _cfb = cfb;
                _inner = inner;
            }

            public override bool CanRead => _inner.CanRead;

            public override bool CanSeek => _inner.CanSeek;

            public override bool CanWrite => _inner.CanWrite;

            public override long Length => _inner.Length;

            public override long Position
            {
                get => _inner.Position;
                set => _inner.Position = value;
            }

            public override void Flush()
            {
                _inner.Flush();
            }

            public override int Read(byte[] buffer, int offset, int count)
            {
                return _inner.Read(buffer, offset, count);
            }

            public override int Read(Span<byte> buffer)
            {
                return _inner.Read(buffer);
            }

            public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
            {
                return _inner.ReadAsync(buffer, offset, count, cancellationToken);
            }

            public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
            {
                return _inner.ReadAsync(buffer, cancellationToken);
            }

            public override long Seek(long offset, SeekOrigin origin)
            {
                return _inner.Seek(offset, origin);
            }

            public override void SetLength(long value)
            {
                _inner.SetLength(value);
            }

            public override void Write(byte[] buffer, int offset, int count)
            {
                _inner.Write(buffer, offset, count);
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing && !_disposed)
                {
                    _disposed = true;
                    _inner.Dispose();
                    _cfb.Dispose();
                }
                base.Dispose(disposing);
            }
        }
    }
}
