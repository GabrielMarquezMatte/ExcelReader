using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Crypto
{
    // An encrypted OOXML file is a CFB container holding EncryptionInfo + EncryptedPackage, not a
    // ZIP. This turns one into a seekable plaintext-ZIP stream that XlsxReader/XlsbReader consume
    // unchanged, and answers the one question Excel.cs's detection needs before it can tell an
    // encrypted OOXML workbook apart from a legacy .xls: does this CFB container hold an
    // "EncryptedPackage" stream at all?
    internal static class EncryptedPackageOpener
    {
        private const long MaxEncryptionInfoBytes = 64 * 1024;

        // Used by detection only — probes the directory, then always rewinds. `ExcelReaderOptions.Default`
        // is deliberately used here rather than the caller's options: this is a structural yes/no
        // question (does an "EncryptedPackage" entry exist?), not a decrypt, so none of the caller's
        // limits/password apply yet.
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
                // Not a well-formed CFB at all: let the XLS path produce its own diagnosis.
                return false;
            }
            finally
            {
                seekableSource.Position = start;
            }
        }

        // Turns an encrypted OOXML CFB container into a seekable, read-only plaintext-ZIP stream.
        // `leaveOpen` follows the same contract Excel.cs's other Open overloads use: when false, the
        // returned stream (and everything it wraps, including `source`) is owned by the caller and
        // must be disposed exactly once — on success via the returned stream's Dispose, on failure
        // here.
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
                // DecryptedPackageStream only takes ownership of the "EncryptedPackage" stream view it
                // opened off `cfb` (see CfbContainer.OpenStreamView) — it never stores `cfb` itself, and
                // its lazy per-segment reads keep seeking back into `cfb.Source` for as long as it's
                // alive. So `cfb` has to outlive it, and something has to dispose both together: that's
                // what OwnedDecryptedStream is for.
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

            [System.Diagnostics.CodeAnalysis.SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
                Justification = "OwnedDecryptedStream owns both _inner and _cfb for the duration of the decrypted read; see EncryptedPackageOpener.Decrypt.")]
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
