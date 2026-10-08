using System.Buffers;

namespace ExcelReader.Core.Reader.Zip.Inflate
{
    /// <summary>
    /// Read-only stream that inflates a raw DEFLATE stream. Owns <c>inner</c> and disposes it.
    /// </summary>
    internal sealed class InflateStream : Stream
    {
        private const int InputSize = 64 * 1024;
        private const int History = 32 * 1024;
        private const int ChunkSize = 64 * 1024;
        private const int OutLimit = History + ChunkSize;
        private const int CanaryStart = OutLimit + InflateDecoder.OutputSlack;
        private const int CanarySize = 16;
        private const byte CanaryByte = 0xA5;

        private readonly Stream _inner;
        private readonly InflateDecoder _decoder = new();
        private readonly byte[] _input = ArrayPool<byte>.Shared.Rent(InputSize + InflateDecoder.InputPadding);
        private readonly byte[] _window = ArrayPool<byte>.Shared.Rent(CanaryStart + CanarySize);
        private int _inPos;
        private int _inEnd;
        private bool _inputEnded;
        private int _outPos;
        private int _readPos;
        private bool _finished;
        private bool _disposed;

        internal InflateStream(Stream inner)
        {
            _inner = inner;
            _window.AsSpan(CanaryStart, CanarySize).Fill(CanaryByte);
        }

        public override bool CanRead => !_disposed;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return Read(buffer.AsSpan(offset, count));
        }

        public override int Read(Span<byte> buffer)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            int total = 0;
            while (total < buffer.Length)
            {
                int copied = CopyOut(buffer[total..]);
                if (copied > 0)
                {
                    total += copied;
                    continue;
                }
                if (_finished)
                {
                    break;
                }
                if (!Advance())
                {
                    FillInput();
                }
            }
            return total;
        }

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            int total = 0;
            while (total < buffer.Length)
            {
                int copied = CopyOut(buffer.Span[total..]);
                if (copied > 0)
                {
                    total += copied;
                    continue;
                }
                if (_finished)
                {
                    break;
                }
                if (!Advance())
                {
                    await FillInputAsync(cancellationToken).ConfigureAwait(false);
                }
            }
            return total;
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        }

        public override void Flush()
        {
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

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                _disposed = true;
                ArrayPool<byte>.Shared.Return(_input);
                ArrayPool<byte>.Shared.Return(_window);
                _decoder.Dispose();
                if (disposing)
                {
                    _inner.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        private int CopyOut(Span<byte> destination)
        {
            int count = Math.Min(_outPos - _readPos, destination.Length);
            _window.AsSpan(_readPos, count).CopyTo(destination);
            _readPos += count;
            return count;
        }

        // Returns false when the decoder needs more input. Only called once everything decoded so far
        // has been copied out, which is what makes sliding the window safe.
        private bool Advance()
        {
            if (_outPos >= OutLimit)
            {
                _window.AsSpan(_outPos - History, History).CopyTo(_window);
                _outPos = History;
                _readPos = History;
            }
            InflateStatus status = _decoder.Decode(_input, ref _inPos, _inEnd, _inputEnded, _window, ref _outPos, OutLimit);
            if (_window.AsSpan(CanaryStart, CanarySize).ContainsAnyExcept(CanaryByte))
            {
                throw new InvalidOperationException("The inflate decoder wrote past its output window.");
            }
            _finished = status == InflateStatus.Finished;
            return status != InflateStatus.NeedInput;
        }

        private void FillInput()
        {
            Compact();
            while (_inEnd < InputSize)
            {
                int read = _inner.Read(_input, _inEnd, InputSize - _inEnd);
                if (read == 0)
                {
                    EndInput();
                    return;
                }
                _inEnd += read;
            }
        }

        private async ValueTask FillInputAsync(CancellationToken cancellationToken)
        {
            Compact();
            while (_inEnd < InputSize)
            {
                int read = await _inner.ReadAsync(_input.AsMemory(_inEnd, InputSize - _inEnd), cancellationToken).ConfigureAwait(false);
                if (read == 0)
                {
                    EndInput();
                    return;
                }
                _inEnd += read;
            }
        }

        private void Compact()
        {
            int remaining = _inEnd - _inPos;
            _input.AsSpan(_inPos, remaining).CopyTo(_input);
            _inPos = 0;
            _inEnd = remaining;
        }

        private void EndInput()
        {
            _inputEnded = true;
            _input.AsSpan(_inEnd, InflateDecoder.InputPadding).Clear();
        }
    }
}
