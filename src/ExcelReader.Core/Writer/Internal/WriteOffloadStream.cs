using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.ExceptionServices;
using System.Threading.Channels;

namespace ExcelReader.Core.Writer.Internal
{
    // Overlaps a ZIP entry's deflate (background thread) with row-serialization (calling thread) for a
    // single entry — the write-side mirror of Reader.PrefetchStream. Opt-in only. Wraps outermost, right
    // on the freshly opened entry stream: that stream's own Write/WriteAsync performs the actual deflate
    // synchronously (DeflateStream under a ZipArchiveEntry), so moving that call onto a background
    // thread lets the caller keep building the next batch of rows concurrently instead of blocking on
    // compression.
    //
    // The ordinary Write/WriteAsync overrides copy into a pooled buffer before returning, since a
    // generic Stream caller may reuse its source span/memory immediately afterward. XlsxSheetWriter/
    // XlsbSheetWriter instead call EnqueueOwned(Async): their row/record buffer is a BiffBuffer, whose
    // Detach() hands over its backing array and rents a fresh one in the same call — so those callers
    // transfer ownership with no copy at all, mirroring how PrefetchStream's read side hands each
    // decompressed chunk to the consumer without copying it either.
    internal sealed class WriteOffloadStream : Stream
    {
        private const int ChannelCapacity = 4;

        private readonly Stream _inner;
        // Owned=true means the buffer came from BiffBuffer.Detach (its own dedicated pool) and must
        // be returned via BiffBuffer.ReturnDetached; Owned=false means it was rented here from
        // ArrayPool<byte>.Shared for the ordinary copying Write/WriteAsync path. Mixing the two
        // without this tag would return an array to the wrong pool.
        private readonly Channel<(byte[] Buffer, int Length, bool Owned)> _channel;
        private readonly Task _consumer;
        private ExceptionDispatchInfo? _consumerException;
        private bool _writerCompleted;
        // Stream.DisposeAsync() forwards to Dispose(bool), so the async path would otherwise re-enter
        // teardown after the consumer task already finished and _inner is already disposed.
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

        // The blocking path here is deliberate, mirroring PrefetchStream's sync Read: it only ever
        // blocks on the bounded channel (backpressure from a slow consumer), never on real I/O, since
        // the actual write/deflate happens on the consumer thread.
        [SuppressMessage("VisualStudio.Threading", "VSTHRD002:Avoid problematic synchronous waits",
            Justification = "Sync Write must block on the bounded channel by design; it never blocks on I/O.")]
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

        // Zero-copy counterpart to Write: takes ownership of `buffer` (obtained from
        // BiffBuffer.Detach) instead of renting a fresh ArrayPool<byte>.Shared array and copying
        // into it. The consumer thread returns it via BiffBuffer.ReturnDetached once written.
        // Caller must never touch `buffer` again after this call returns.
        internal void EnqueueOwned(byte[] buffer, int length)
        {
            ThrowIfFaulted();
            if (length == 0)
            {
                BiffBuffer.ReturnDetached(buffer);
                return;
            }
            EnqueueSync((buffer, length, Owned: true));
        }

        [SuppressMessage("VisualStudio.Threading", "VSTHRD002:Avoid problematic synchronous waits",
            Justification = "Sync Write must block on the bounded channel by design; it never blocks on I/O.")]
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
            catch (ChannelClosedException)
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
            catch (ChannelClosedException)
            {
                ArrayPool<byte>.Shared.Return(rented);
                ThrowIfFaulted();
                throw;
            }
            ThrowIfFaulted();
        }

        // Async, zero-copy counterpart to EnqueueOwned. See EnqueueOwned's remarks.
        internal async ValueTask EnqueueOwnedAsync(byte[] buffer, int length, CancellationToken cancellationToken = default)
        {
            ThrowIfFaulted();
            if (length == 0)
            {
                BiffBuffer.ReturnDetached(buffer);
                return;
            }
            try
            {
                await _channel.Writer.WriteAsync((buffer, length, Owned: true), cancellationToken).ConfigureAwait(false);
            }
            catch (ChannelClosedException)
            {
                BiffBuffer.ReturnDetached(buffer);
                ThrowIfFaulted();
                throw;
            }
            ThrowIfFaulted();
        }

        [SuppressMessage("VisualStudio.Threading", "VSTHRD002:Avoid problematic synchronous waits",
            Justification = "Sync Flush must block until the background writer catches up; every caller of this path (XlsxSheetWriter's sync row API) already accepts blocking on I/O.")]
        public override void Flush()
        {
            FlushAsync(CancellationToken.None).GetAwaiter().GetResult();
        }

        // Drains every enqueued write through the consumer before returning, so a caller that flushes
        // right before disposing the real entry stream (see XlsxSheetWriter.EndAsync) never closes it
        // with buffered bytes still in flight. Safe to call at most once per stream: every real call
        // site here flushes only immediately before Dispose/DisposeAsync, never writes afterward.
        [SuppressMessage("VisualStudio.Threading", "VSTHRD003:Avoid awaiting foreign Tasks",
            Justification = "_consumer is this instance's own consumer loop, started in the constructor and always brought to completion exactly once, here or in Dispose(Async).")]
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

        // Runs on a pooled thread pool thread for the entry's whole lifetime. Every exception path is
        // caught so the Task itself always completes successfully — Dispose/DisposeAsync/FlushAsync
        // await it without risking an unobserved-exception or a rethrow at a moment nobody is prepared
        // to catch it. On failure, completes the channel so a producer already blocked on a full
        // channel (waiting for room the now-dead consumer will never free) unblocks instead of hanging.
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
            }
        }

        [SuppressMessage("VisualStudio.Threading", "VSTHRD002:Avoid problematic synchronous waits",
            Justification = "Dispose must not return before the consumer thread stops touching _inner; ConsumeAsync catches every exception, so this never blocks long or throws.")]
        protected override void Dispose(bool disposing)
        {
            if (disposing && !_disposed)
            {
                _disposed = true;
                CompleteWriterOnce();
                _consumer.GetAwaiter().GetResult();
                ThrowIfFaulted();
                _inner.Dispose();
            }
            base.Dispose(disposing);
        }

        [SuppressMessage("VisualStudio.Threading", "VSTHRD003:Avoid awaiting foreign Tasks",
            Justification = "_consumer is this instance's own consumer loop, started in the constructor and always brought to completion exactly once, here or in Dispose(bool).")]
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
            ThrowIfFaulted();
            await _inner.DisposeAsync().ConfigureAwait(false);
            await base.DisposeAsync().ConfigureAwait(false);
        }
    }
}
