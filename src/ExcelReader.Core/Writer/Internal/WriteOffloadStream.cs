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
    // Every write is copied into a pooled buffer before this returns, since the caller (XlsxSheetWriter/
    // XlsbSheetWriter) reuses its own row buffer immediately after handing bytes here — unlike
    // PrefetchStream's read side, where each chunk is a fresh buffer the consumer takes ownership of.
    internal sealed class WriteOffloadStream : Stream
    {
        private const int ChannelCapacity = 4;

        private readonly Stream _inner;
        private readonly Channel<(byte[] Buffer, int Length)> _channel;
        private readonly Task _consumer;
        private ExceptionDispatchInfo? _consumerException;
        private bool _writerCompleted;
        // Stream.DisposeAsync() forwards to Dispose(bool), so the async path would otherwise re-enter
        // teardown after the consumer task already finished and _inner is already disposed.
        private bool _disposed;

        internal WriteOffloadStream(Stream inner)
        {
            _inner = inner;
            _channel = Channel.CreateBounded<(byte[] Buffer, int Length)>(new BoundedChannelOptions(ChannelCapacity)
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
            EnqueueSync((rented, buffer.Length));
        }

        [SuppressMessage("VisualStudio.Threading", "VSTHRD002:Avoid problematic synchronous waits",
            Justification = "Sync Write must block on the bounded channel by design; it never blocks on I/O.")]
        private void EnqueueSync((byte[] Buffer, int Length) item)
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
                ArrayPool<byte>.Shared.Return(item.Buffer);
                ThrowIfFaulted();
                throw;
            }
            ThrowIfFaulted();
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
                await _channel.Writer.WriteAsync((rented, buffer.Length), cancellationToken).ConfigureAwait(false);
            }
            catch (ChannelClosedException)
            {
                ArrayPool<byte>.Shared.Return(rented);
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
                await foreach ((byte[] buf, int len) in _channel.Reader.ReadAllAsync().ConfigureAwait(false))
                {
                    try
                    {
                        await _inner.WriteAsync(buf.AsMemory(0, len)).ConfigureAwait(false);
                    }
                    finally
                    {
                        ArrayPool<byte>.Shared.Return(buf);
                    }
                }
            }
            catch (Exception ex)
            {
                _consumerException = ExceptionDispatchInfo.Capture(ex);
                _channel.Writer.TryComplete();
            }
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "WriteOffloadStream owns the wrapped entry stream for the duration of the decorated write.")]
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

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "WriteOffloadStream owns the wrapped entry stream for the duration of the decorated write.")]
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
