using System.Buffers;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Reader.Csv
{
    /// <summary>
    /// Serves a forward-only stream to one pass at a time, and lets a later pass start over for as long
    /// as everything read so far still fits in <c>limit</c> bytes. Disposing ends a pass;
    /// <see cref="Release"/> closes the stream underneath.
    /// </summary>
    internal sealed class ReplayStream(Stream inner, bool leaveOpen, int limit) : ReadOnlyStream
    {
        private byte[]? _recorded = [];
        private int _recordedLength;
        private int _position;
        private int _busy;

        internal bool TryBeginPass()
        {
            if (Interlocked.CompareExchange(ref _busy, 1, 0) != 0)
            {
                return false;
            }
            if (_recorded is null)
            {
                Volatile.Write(ref _busy, 0);
                return false;
            }
            _position = 0;
            return true;
        }

        internal void Release()
        {
            DropRecording();
            if (!leaveOpen)
            {
                inner.Dispose();
            }
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return Read(buffer.AsSpan(offset, count));
        }

        public override int Read(Span<byte> buffer)
        {
            if (TryReplay(buffer, out int replayed))
            {
                return replayed;
            }
            int read = inner.Read(buffer);
            Record(buffer[..read]);
            return read;
        }

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            if (TryReplay(buffer.Span, out int replayed))
            {
                return replayed;
            }
            int read = await inner.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
            Record(buffer.Span[..read]);
            return read;
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        }

        private bool TryReplay(Span<byte> buffer, out int count)
        {
            count = 0;
            if (_recorded is not { } recorded || _position >= _recordedLength || buffer.IsEmpty)
            {
                return false;
            }
            count = Math.Min(buffer.Length, _recordedLength - _position);
            recorded.AsSpan(_position, count).CopyTo(buffer);
            _position += count;
            return true;
        }

        private void Record(ReadOnlySpan<byte> data)
        {
            if (_recorded is not { } recorded)
            {
                return;
            }
            int needed = _recordedLength + data.Length;
            if (needed > limit)
            {
                DropRecording();
                return;
            }
            if (needed > recorded.Length)
            {
                byte[] bigger = ArrayPool<byte>.Shared.Rent(Math.Max(needed, Math.Max(64 * 1024, recorded.Length * 2)));
                recorded.AsSpan(0, _recordedLength).CopyTo(bigger);
                DropRecording();
                _recorded = recorded = bigger;
            }
            data.CopyTo(recorded.AsSpan(_recordedLength));
            _recordedLength = needed;
            _position += data.Length;
        }

        private void DropRecording()
        {
            if (_recorded is { Length: > 0 } recorded)
            {
                ArrayPool<byte>.Shared.Return(recorded);
            }
            _recorded = null;
        }

        protected override void Dispose(bool disposing)
        {
            Volatile.Write(ref _busy, 0);
            base.Dispose(disposing);
        }
    }
}
