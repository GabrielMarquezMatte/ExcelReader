using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Zip
{
    /// <summary>
    /// The write-only stream of one <see cref="ZipPackageWriter"/> entry. It checksums and compresses
    /// what it is given; closing it completes the local header in place, by seeking back, or with a
    /// data descriptor.
    /// </summary>
    internal sealed class ZipEntryStream : Stream
    {
        private const int FlushThreshold = 64 * 1024;

        private readonly ZipPackageWriter _owner;
        private readonly ZipEntryRecord _record;
        [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed",
            Justification = "Released by Release(), which every close path and Abandon call.")]
        private readonly BiffBuffer _pending = new(FlushThreshold * 2);
        [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed",
            Justification = "Released by Release(), which every close path and Abandon call.")]
        private readonly DeflateEncoder? _encoder;
        [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed",
            Justification = "Finished by FinishCompression and released by Release().")]
        private readonly DeflateStream? _zlib;
        private uint _crc;
        private long _uncompressed;
        private long _compressed;
        private bool _headerFlushed;
        private bool _closed;

        internal ZipEntryStream(ZipPackageWriter owner, ZipEntryRecord record)
        {
            _owner = owner;
            _record = record;
            if (!owner.CanSeek)
            {
                record.Flags |= ZipRecords.FlagDataDescriptor;
            }
            ZipRecords.WriteLocalHeader(_pending, record);
            if (record.Level == CompressionLevel.Fastest)
            {
                _encoder = new DeflateEncoder();
            }
            else if (record.Method == ZipRecords.MethodDeflate)
            {
                _zlib = new DeflateStream(new BufferSink(_pending), record.Level);
            }
        }

        private enum Completion
        {
            Done,
            PatchHeader,
            PatchZip64,
        }

        public override bool CanRead
        {
            get
            {
                return false;
            }
        }

        public override bool CanSeek
        {
            get
            {
                return false;
            }
        }

        public override bool CanWrite
        {
            get
            {
                return !_closed;
            }
        }

        public override long Length
        {
            get
            {
                throw new NotSupportedException();
            }
        }

        public override long Position
        {
            get
            {
                throw new NotSupportedException();
            }
            set
            {
                throw new NotSupportedException();
            }
        }

        public override void Flush()
        {
            if (_closed)
            {
                return;
            }
            try
            {
                _owner.FlushOutput();
            }
            catch
            {
                Fail();
                throw;
            }
        }

        public override async Task FlushAsync(CancellationToken cancellationToken)
        {
            if (_closed)
            {
                return;
            }
            try
            {
                await _owner.FlushOutputAsync(cancellationToken).ConfigureAwait(false);
            }
            catch
            {
                Fail();
                throw;
            }
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotSupportedException();
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            Write(buffer.AsSpan(offset, count));
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            ObjectDisposedException.ThrowIf(_closed, this);
            try
            {
                Accept(buffer);
                if (_pending.Length >= FlushThreshold)
                {
                    FlushPending();
                }
            }
            catch
            {
                Fail();
                throw;
            }
        }

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return WriteAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        }

        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
        {
            ObjectDisposedException.ThrowIf(_closed, this);
            try
            {
                Accept(buffer.Span);
                if (_pending.Length >= FlushThreshold)
                {
                    await FlushPendingAsync(cancellationToken).ConfigureAwait(false);
                }
            }
            catch
            {
                Fail();
                throw;
            }
        }

        internal void Abandon()
        {
            if (_closed)
            {
                return;
            }
            _closed = true;
            Release();
        }

        // A failed write leaves the output position unknown, so the entry and the archive are dead.
        private void Fail()
        {
            _closed = true;
            _owner.EntryFailed();
            Release();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && !_closed)
            {
                _closed = true;
                try
                {
                    Completion completion = PrepareCompletion();
                    FlushPending();
                    CompleteAfterFlush(completion);
                    _owner.EntryClosed(_record);
                }
                catch
                {
                    _owner.EntryFailed();
                    throw;
                }
                finally
                {
                    Release();
                }
            }
            base.Dispose(disposing);
        }

        public override async ValueTask DisposeAsync()
        {
            if (!_closed)
            {
                _closed = true;
                try
                {
                    Completion completion = PrepareCompletion();
                    await FlushPendingAsync(CancellationToken.None).ConfigureAwait(false);
                    await CompleteAfterFlushAsync(completion).ConfigureAwait(false);
                    _owner.EntryClosed(_record);
                }
                catch
                {
                    _owner.EntryFailed();
                    throw;
                }
                finally
                {
                    Release();
                }
            }
            await base.DisposeAsync().ConfigureAwait(false);
        }

        private void Accept(ReadOnlySpan<byte> buffer)
        {
            _crc = Crc32.Append(_crc, buffer);
            _uncompressed += buffer.Length;
            if (_encoder is not null)
            {
                _encoder.Write(buffer, _pending);
                return;
            }
            if (_zlib is not null)
            {
                _zlib.Write(buffer);
                return;
            }
            _pending.Write(buffer);
        }

        private Completion PrepareCompletion()
        {
            if (_uncompressed == 0 && !_headerFlushed)
            {
                MarkEmpty();
                return Completion.Done;
            }
            FinishCompression();
            _record.Crc = _crc;
            _record.UncompressedSize = _uncompressed;
            _record.CompressedSize = _compressed + _pending.Length - PendingHeaderLength();
            if ((_record.Flags & ZipRecords.FlagDataDescriptor) != 0)
            {
                ZipRecords.WriteDataDescriptor(_pending, _record);
                return Completion.Done;
            }
            if (!_headerFlushed)
            {
                ZipRecords.CrcAndSizes(_record).CopyTo(_pending.Slice(ZipRecords.LocalCrcOffset, 12));
                return Completion.Done;
            }
            if (_record.NeedsZip64Sizes)
            {
                _record.Flags |= ZipRecords.FlagDataDescriptor;
                return Completion.PatchZip64;
            }
            return Completion.PatchHeader;
        }

        private void CompleteAfterFlush(Completion completion)
        {
            if (completion == Completion.PatchHeader)
            {
                _owner.Patch(_record.Offset + ZipRecords.LocalCrcOffset, ZipRecords.CrcAndSizes(_record));
                return;
            }
            if (completion == Completion.PatchZip64)
            {
                _owner.Patch(_record.Offset + ZipRecords.LocalVersionOffset, ZipRecords.VersionAndFlags(_record));
                ZipRecords.WriteDataDescriptor(_pending, _record);
                FlushPending();
            }
        }

        private async ValueTask CompleteAfterFlushAsync(Completion completion)
        {
            if (completion == Completion.PatchHeader)
            {
                await _owner.PatchAsync(_record.Offset + ZipRecords.LocalCrcOffset, ZipRecords.CrcAndSizes(_record), CancellationToken.None).ConfigureAwait(false);
                return;
            }
            if (completion == Completion.PatchZip64)
            {
                await _owner.PatchAsync(_record.Offset + ZipRecords.LocalVersionOffset, ZipRecords.VersionAndFlags(_record), CancellationToken.None).ConfigureAwait(false);
                ZipRecords.WriteDataDescriptor(_pending, _record);
                await FlushPendingAsync(CancellationToken.None).ConfigureAwait(false);
            }
        }

        // ZipArchive writes an entry that received no data as stored, with no data descriptor.
        private void MarkEmpty()
        {
            _record.Method = ZipRecords.MethodStored;
            _record.Flags = (ushort)(_record.Flags & ~ZipRecords.FlagDataDescriptor);
            _pending.Reset();
            ZipRecords.WriteLocalHeader(_pending, _record);
        }

        private void FinishCompression()
        {
            _encoder?.Finish(_pending);
            _zlib?.Dispose();
        }

        private int PendingHeaderLength()
        {
            return _headerFlushed ? 0 : _record.LocalHeaderLength;
        }

        private void FlushPending()
        {
            _owner.WriteRaw(_pending.Span);
            CountFlushed();
        }

        private async ValueTask FlushPendingAsync(CancellationToken ct)
        {
            await _owner.WriteRawAsync(_pending.Memory, ct).ConfigureAwait(false);
            CountFlushed();
        }

        private void CountFlushed()
        {
            _compressed += _pending.Length - PendingHeaderLength();
            _headerFlushed = true;
            _pending.Reset();
        }

        private void Release()
        {
            _encoder?.Dispose();
            FailureCleanup.Dispose(_zlib);
            _pending.Dispose();
        }

        private sealed class BufferSink : Stream
        {
            private readonly BiffBuffer _target;

            internal BufferSink(BiffBuffer target)
            {
                _target = target;
            }

            public override bool CanRead
            {
                get
                {
                    return false;
                }
            }

            public override bool CanSeek
            {
                get
                {
                    return false;
                }
            }

            public override bool CanWrite
            {
                get
                {
                    return true;
                }
            }

            public override long Length
            {
                get
                {
                    throw new NotSupportedException();
                }
            }

            public override long Position
            {
                get
                {
                    throw new NotSupportedException();
                }
                set
                {
                    throw new NotSupportedException();
                }
            }

            public override void Flush()
            {
            }

            public override int Read(byte[] buffer, int offset, int count)
            {
                throw new NotSupportedException();
            }

            public override long Seek(long offset, SeekOrigin origin)
            {
                throw new NotSupportedException();
            }

            public override void SetLength(long value)
            {
                throw new NotSupportedException();
            }

            public override void Write(byte[] buffer, int offset, int count)
            {
                _target.Write(buffer.AsSpan(offset, count));
            }

            public override void Write(ReadOnlySpan<byte> buffer)
            {
                _target.Write(buffer);
            }
        }
    }
}
