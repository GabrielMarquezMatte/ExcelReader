using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.ExceptionServices;
using System.Threading.Channels;

namespace ExcelReader.Core.Reader
{
    // Overlaps zlib inflate (producer thread) with XML/record parsing (consumer thread) for a single
    // ZIP entry. Opt-in only (see ExcelReaderOptions.PrefetchDecompression) — wraps innermost, under
    // LimitedReadStream, so the decompressed-byte counters stay on the consumer thread untouched.
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
        // Stream.DisposeAsync() forwards to Dispose(bool), so the async path would otherwise re-enter
        // teardown after _cts is already disposed and Cancel() would throw ObjectDisposedException.
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

        // Read-only decorator: nothing is ever written, and _inner is touched exclusively by the
        // producer task, so there is nothing for a consumer-thread Flush to do or safely reach.
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
            return ConsumeCurrentChunk(buffer);
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
            return ConsumeCurrentChunk(buffer.Span);
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

        // The blocking GetAwaiter().GetResult() below is the sync Read path's whole point: it lets a
        // caller on the ordinary synchronous path still benefit from the background inflate without
        // an async rewrite. It only ever blocks on the bounded channel, never on I/O.
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

        // Isolated so the ValueTask from WaitToReadAsync is consumed along exactly one path (either
        // the already-completed branch or the GetResult() branch), never both.
        [SuppressMessage("VisualStudio.Threading", "VSTHRD002:Avoid problematic synchronous waits",
            Justification = "Sync Read must block on the producer channel by design; it never blocks on I/O.")]
        private bool WaitForNextChunkSync()
        {
            ValueTask<bool> waitTask = _channel.Reader.WaitToReadAsync();
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

        // Copies as much of the current chunk as fits, returning the buffer to the pool only once the
        // caller has drained it fully — a caller asking for fewer bytes than a chunk holds must see the
        // remainder on its next call, not lose it.
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

        // Runs on a pooled thread pool thread for the entry's whole lifetime. Blocking Read here is
        // deliberate: this is CPU-bound inflate, not blocking I/O, so occupying the thread is not a
        // starvation bug (see docs/parallel-prefetch.md). Every exception path is caught so the Task
        // itself always completes successfully — Dispose/DisposeAsync await it without risking an
        // unobserved-exception or a rethrow at a moment nobody is prepared to catch it.
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
                    pending = null; // ownership passed to the channel/consumer
                }
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested)
            {
                // Early abandonment (consumer disposed while WriteAsync was blocked on a full channel).
                // Not an error — the consumer already stopped reading.
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

        // Returns every pooled buffer still sitting in the channel plus the partially-consumed current
        // one. Runs after the producer task has finished, so no new items can arrive underneath it.
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

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "PrefetchStream owns the wrapped entry stream for the duration of the decorated read.")]
        [SuppressMessage("VisualStudio.Threading", "VSTHRD002:Avoid problematic synchronous waits",
            Justification = "Dispose must not return before the producer thread stops touching _inner; ProduceAsync catches every exception, so this never blocks long or throws.")]
        [SuppressMessage("Reliability", "CA1849:Call async methods when in an async method",
            Justification = "This is the synchronous Dispose(bool) override — there is no async context to await CancelAsync from.")]
        [SuppressMessage("VisualStudio.Threading", "VSTHRD103:Call async methods when in an async method",
            Justification = "This is the synchronous Dispose(bool) override — there is no async context to await CancelAsync from.")]
        [SuppressMessage("Major Code Smell", "S6966:Awaitable method should be used",
            Justification = "This is the synchronous Dispose(bool) override — there is no async context to await CancelAsync from.")]
        protected override void Dispose(bool disposing)
        {
            if (disposing && !_disposed)
            {
                _disposed = true;
                _cts.Cancel();
                _producer.GetAwaiter().GetResult();
                DrainRemainingBuffers();
                _cts.Dispose();
                _inner.Dispose();
            }
            base.Dispose(disposing);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "PrefetchStream owns the wrapped entry stream for the duration of the decorated read.")]
        [SuppressMessage("VisualStudio.Threading", "VSTHRD003:Avoid awaiting foreign Tasks",
            Justification = "_producer is this instance's own producer loop, started in the constructor and always brought to completion exactly once, here or in Dispose(bool) — not a fire-and-forget foreign Task.")]
        public override async ValueTask DisposeAsync()
        {
            if (!_disposed)
            {
                _disposed = true;
                await _cts.CancelAsync().ConfigureAwait(false);
                await _producer.ConfigureAwait(false);
                DrainRemainingBuffers();
                _cts.Dispose();
                await _inner.DisposeAsync().ConfigureAwait(false);
            }
            await base.DisposeAsync().ConfigureAwait(false);
        }
    }
}
