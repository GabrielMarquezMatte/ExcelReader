using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.ExceptionServices;
using System.Threading.Channels;

namespace ExcelReader.Core.Writer.Internal
{
    internal sealed class WriteOffloadStream : Stream
    {
        private const int ChannelCapacity = 4;

        private readonly Stream _inner;
        private readonly Channel<(byte[] Buffer, int Length, bool Owned)> _channel;
        private readonly Task _consumer;
        private ExceptionDispatchInfo? _consumerException;
        private bool _writerCompleted;
        private bool _disposed;

        internal WriteOffloadStream(Stream inner)
        {
            _inner = inner;
            _channel = Channel.CreateBounded<(byte[] Buffer, int Length, bool Owned)>(new BoundedChannelOptions(ChannelCapacity)
            {
                SingleReader = true,
                SingleWriter = true,
                FullMode = BoundedChannelFullMode.Wait,
            });
            _consumer = Task.Run(ConsumeAsync);
        }

        public override bool CanRead => false;

        public override bool CanSeek => false;

        public override bool CanWrite => true;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException(); set => throw new NotSupportedException();
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
            ThrowIfFaulted();
            if (buffer.IsEmpty)
            {
                return;
            }
            byte[] rented = ArrayPool<byte>.Shared.Rent(buffer.Length);
            buffer.CopyTo(rented);
            EnqueueSync((rented, buffer.Length, Owned: false));
        }

        internal void EnqueueOwned(byte[] buffer, int length)
        {
            if (length == 0 || _consumerException is not null)
            {
                BiffBuffer.ReturnDetached(buffer);
                ThrowIfFaulted();
                return;
            }
            EnqueueSync((buffer, length, Owned: true));
        }

        private void EnqueueSync((byte[] Buffer, int Length, bool Owned) item)
        {
            try
            {
                ValueTask writeTask = _channel.Writer.WriteAsync(item);
                if (!writeTask.IsCompletedSuccessfully)
                {
                    writeTask.AsTask().GetAwaiter().GetResult();
                }
            }
            catch
            {
                ReturnBuffer(item.Buffer, item.Owned);
                ThrowIfFaulted();
                throw;
            }
            ThrowIfFaulted();
        }

        private static void ReturnBuffer(byte[] buffer, bool owned)
        {
            if (owned)
            {
                BiffBuffer.ReturnDetached(buffer);
                return;
            }
            ArrayPool<byte>.Shared.Return(buffer);
        }

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return WriteAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        }

        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
        {
            ThrowIfFaulted();
            if (buffer.IsEmpty)
            {
                return;
            }
            byte[] rented = ArrayPool<byte>.Shared.Rent(buffer.Length);
            buffer.Span.CopyTo(rented);
            try
            {
                await _channel.Writer.WriteAsync((rented, buffer.Length, Owned: false), cancellationToken).ConfigureAwait(false);
            }
            catch
            {
                ArrayPool<byte>.Shared.Return(rented);
                ThrowIfFaulted();
                throw;
            }
            ThrowIfFaulted();
        }

        internal async ValueTask EnqueueOwnedAsync(byte[] buffer, int length, CancellationToken cancellationToken = default)
        {
            if (length == 0 || _consumerException is not null)
            {
                BiffBuffer.ReturnDetached(buffer);
                ThrowIfFaulted();
                return;
            }
            try
            {
                await _channel.Writer.WriteAsync((buffer, length, Owned: true), cancellationToken).ConfigureAwait(false);
            }
            catch
            {
                BiffBuffer.ReturnDetached(buffer);
                ThrowIfFaulted();
                throw;
            }
            ThrowIfFaulted();
        }

        public override void Flush()
        {
            FlushAsync(CancellationToken.None).GetAwaiter().GetResult();
        }

        public override async Task FlushAsync(CancellationToken cancellationToken)
        {
            CompleteWriterOnce();
            await _consumer.WaitAsync(cancellationToken).ConfigureAwait(false);
            ThrowIfFaulted();
            await _inner.FlushAsync(cancellationToken).ConfigureAwait(false);
        }

        private void CompleteWriterOnce()
        {
            if (!_writerCompleted)
            {
                _writerCompleted = true;
                _channel.Writer.TryComplete();
            }
        }

        private void ThrowIfFaulted()
        {
            _consumerException?.Throw();
        }

        [SuppressMessage("Design", "CA1031:Do not catch general exception types",
            Justification = "Any exception here (a real I/O failure, a limit exceeded downstream) must reach the producer's next Write/WriteAsync/Flush with its original type preserved, so it is captured via ExceptionDispatchInfo rather than left to fault this Task unobserved.")]
        private async Task ConsumeAsync()
        {
            try
            {
                await foreach ((byte[] buf, int len, bool owned) in _channel.Reader.ReadAllAsync().ConfigureAwait(false))
                {
                    try
                    {
                        await _inner.WriteAsync(buf.AsMemory(0, len)).ConfigureAwait(false);
                    }
                    finally
                    {
                        ReturnBuffer(buf, owned);
                    }
                }
            }
            catch (Exception ex)
            {
                _consumerException = ExceptionDispatchInfo.Capture(ex);
                _channel.Writer.TryComplete();
                while (_channel.Reader.TryRead(out (byte[] Buffer, int Length, bool Owned) queued))
                {
                    ReturnBuffer(queued.Buffer, queued.Owned);
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && !_disposed)
            {
                _disposed = true;
                CompleteWriterOnce();
                _consumer.GetAwaiter().GetResult();
                if (_consumerException is not null)
                {
                    FailureCleanup.Dispose(_inner);
                    _consumerException.Throw();
                }
                _inner.Dispose();
            }
            base.Dispose(disposing);
        }

        public override async ValueTask DisposeAsync()
        {
            if (_disposed)
            {
                await base.DisposeAsync().ConfigureAwait(false);
                return;
            }
            _disposed = true;
            CompleteWriterOnce();
            await _consumer.ConfigureAwait(false);
            if (_consumerException is not null)
            {
                await FailureCleanup.DisposeAsync(_inner).ConfigureAwait(false);
                _consumerException.Throw();
            }
            await _inner.DisposeAsync().ConfigureAwait(false);
            await base.DisposeAsync().ConfigureAwait(false);
        }
    }
}
