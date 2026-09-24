using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.ExceptionServices;
using System.Threading.Channels;

namespace ExcelReader.Core.Reader.Zip
{
    internal sealed class PrefetchStream : Stream
    {
        private const int ChunkSize = 64 * 1024;
        private const int ChannelCapacity = 4;

        private readonly Stream _inner;
        private readonly Channel<(byte[] Buffer, int Length)> _channel;
        private readonly CancellationTokenSource _cts;
        private readonly Task _producer;
        private byte[]? _currentBuffer;
        private int _currentLength;
        private int _currentOffset;
        private ExceptionDispatchInfo? _producerException;
        private bool _disposed;

        internal PrefetchStream(Stream inner)
        {
            _inner = inner;
            _channel = Channel.CreateBounded<(byte[] Buffer, int Length)>(new BoundedChannelOptions(ChannelCapacity)
            {
                SingleReader = true,
                SingleWriter = true,
                FullMode = BoundedChannelFullMode.Wait,
            });
            _cts = new CancellationTokenSource();
            _producer = Task.Run(() => ProduceAsync(_cts.Token), _cts.Token);
        }

        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException(); set => throw new NotSupportedException();
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
            if (buffer.Length == 0)
            {
                return 0;
            }
            if (!EnsureCurrentChunk())
            {
                return 0;
            }
            return ConsumeAvailableChunks(buffer);
        }

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            if (buffer.Length == 0)
            {
                return 0;
            }
            if (!await EnsureCurrentChunkAsync(cancellationToken).ConfigureAwait(false))
            {
                return 0;
            }
            return ConsumeAvailableChunks(buffer.Span);
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
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
            throw new NotSupportedException();
        }

        private bool EnsureCurrentChunk()
        {
            if (_currentBuffer is not null)
            {
                return true;
            }
            if (_channel.Reader.TryRead(out var item))
            {
                SetCurrent(item);
                return true;
            }
            if (WaitForNextChunkSync() && _channel.Reader.TryRead(out item))
            {
                SetCurrent(item);
                return true;
            }
            _producerException?.Throw();
            return false;
        }

        private bool WaitForNextChunkSync()
        {
            ValueTask<bool> waitTask = _channel.Reader.WaitToReadAsync(_cts.Token);
            if (waitTask.IsCompletedSuccessfully)
            {
                return waitTask.Result;
            }
            return waitTask.AsTask().GetAwaiter().GetResult();
        }

        private async ValueTask<bool> EnsureCurrentChunkAsync(CancellationToken cancellationToken)
        {
            if (_currentBuffer is not null)
            {
                return true;
            }
            if (_channel.Reader.TryRead(out var item))
            {
                SetCurrent(item);
                return true;
            }
            bool canRead = await _channel.Reader.WaitToReadAsync(cancellationToken).ConfigureAwait(false);
            if (canRead && _channel.Reader.TryRead(out item))
            {
                SetCurrent(item);
                return true;
            }
            _producerException?.Throw();
            return false;
        }

        private void SetCurrent((byte[] Buffer, int Length) item)
        {
            _currentBuffer = item.Buffer;
            _currentLength = item.Length;
            _currentOffset = 0;
        }

        private int ConsumeCurrentChunk(Span<byte> destination)
        {
            byte[] buffer = _currentBuffer!;
            int available = _currentLength - _currentOffset;
            int toCopy = Math.Min(available, destination.Length);
            buffer.AsSpan(_currentOffset, toCopy).CopyTo(destination);
            _currentOffset += toCopy;
            if (_currentOffset >= _currentLength)
            {
                ArrayPool<byte>.Shared.Return(buffer);
                _currentBuffer = null;
            }
            return toCopy;
        }

        private int ConsumeAvailableChunks(Span<byte> destination)
        {
            int total = ConsumeCurrentChunk(destination);
            while (total < destination.Length && _currentBuffer is null && _channel.Reader.TryRead(out var item))
            {
                SetCurrent(item);
                total += ConsumeCurrentChunk(destination[total..]);
            }
            return total;
        }

        [SuppressMessage("Design", "CA1031:Do not catch general exception types",
            Justification = "Any exception here (truncated entry, corrupt deflate, limit exceeded) must reach the consumer's next Read/ReadAsync with its original type preserved, so it is captured via ExceptionDispatchInfo rather than left to fault the producer Task unobserved.")]
        private async Task ProduceAsync(CancellationToken token)
        {
            byte[]? pending = null;
            try
            {
                while (true)
                {
                    pending = ArrayPool<byte>.Shared.Rent(ChunkSize);
                    int read = _inner.Read(pending.AsSpan(0, ChunkSize));
                    if (read <= 0)
                    {
                        break;
                    }
                    await _channel.Writer.WriteAsync((pending, read), token).ConfigureAwait(false);
                    pending = null;
                }
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested)
            {
            }
            catch (Exception ex)
            {
                _producerException = ExceptionDispatchInfo.Capture(ex);
            }
            finally
            {
                if (pending is not null)
                {
                    ArrayPool<byte>.Shared.Return(pending);
                }
                _channel.Writer.TryComplete();
            }
        }

        private void DrainRemainingBuffers()
        {
            while (_channel.Reader.TryRead(out var item))
            {
                ArrayPool<byte>.Shared.Return(item.Buffer);
            }
            if (_currentBuffer is not null)
            {
                ArrayPool<byte>.Shared.Return(_currentBuffer);
                _currentBuffer = null;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && !_disposed)
            {
                _disposed = true;
                _cts.Cancel();
                try
                {
                    _producer.GetAwaiter().GetResult();
                }
                catch (OperationCanceledException) when (_cts.IsCancellationRequested)
                {
                }
                DrainRemainingBuffers();
                _cts.Dispose();
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
            await _cts.CancelAsync().ConfigureAwait(false);
            try
            {
                await _producer.ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (_cts.IsCancellationRequested)
            {
            }
            DrainRemainingBuffers();
            _cts.Dispose();
            await _inner.DisposeAsync().ConfigureAwait(false);
            await base.DisposeAsync().ConfigureAwait(false);
        }
    }
}
