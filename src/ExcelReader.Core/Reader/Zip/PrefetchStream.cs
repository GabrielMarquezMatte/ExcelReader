using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.ExceptionServices;
using System.Threading.Channels;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Writer.Zip;

namespace ExcelReader.Core.Reader.Zip
{
    internal sealed class PrefetchStream : ReadOnlyStream
    {
        private const int ChunkSize = 64 * 1024;
        private const int ChannelCapacity = 4;

        private readonly Stream _inner;
        private readonly Channel<(byte[] Buffer, int Length)> _channel;
        private readonly CancellationTokenSource _cts;
        private readonly Task _producer;
        private readonly long _declaredLength;
        private readonly uint? _expectedCrc;
        private byte[]? _currentBuffer;
        private int _currentLength;
        private int _currentOffset;
        private ExceptionDispatchInfo? _producerException;
        private bool _disposed;

        // With expectedCrc, the producer checksums what it reads and fails instead of handing over the
        // chunk that completes declaredLength, so the check runs off the consumer's thread.
        internal PrefetchStream(Stream inner, long declaredLength = 0, uint? expectedCrc = null)
        {
            _inner = inner;
            _declaredLength = declaredLength;
            _expectedCrc = expectedCrc;
            _channel = Channel.CreateBounded<(byte[] Buffer, int Length)>(new BoundedChannelOptions(ChannelCapacity)
            {
                SingleReader = true,
                SingleWriter = true,
                FullMode = BoundedChannelFullMode.Wait,
            });
            _cts = new CancellationTokenSource();
            _producer = Task.Run(() => ProduceAsync(_cts.Token), _cts.Token);
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
            long produced = 0;
            uint crc = 0;
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
                    // Only the declared bytes are summed, so an entry that inflates past its size still has them checked.
                    if (_expectedCrc is uint expected && produced < _declaredLength)
                    {
                        int counted = (int)Math.Min(read, _declaredLength - produced);
                        crc = Crc32.Append(crc, pending.AsSpan(0, counted));
                        produced += counted;
                        if (produced == _declaredLength && crc != expected)
                        {
                            throw new InvalidDataException(LimitedReadStream.ChecksumMismatch);
                        }
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
