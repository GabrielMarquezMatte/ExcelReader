using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Native.Reading
{
    internal sealed unsafe class CallbackReadStream : ReadOnlyStream
    {
        private readonly delegate* unmanaged<void*, byte*, long, long> _read;
        private readonly delegate* unmanaged<void*, void> _release;
        private readonly void* _userData;
        private ReadOnlyMemory<byte> _peeked;
        private int _released;

        private CallbackReadStream(in NativeStreamRaw raw)
        {
            _read = raw.Read;
            _release = raw.Release;
            _userData = raw.UserData;
        }

        internal static bool TryCreate(
            NativeStreamRaw* raw, bool hasOutHandle, [NotNullWhen(true)] out CallbackReadStream? stream, [NotNullWhen(false)] out string? error)
        {
            stream = null;
            if (raw is null || !hasOutHandle || raw->StructSize != sizeof(NativeStreamRaw) || raw->Read is null)
            {
                error = "xl_open_stream needs a non-NULL stream whose struct_size is sizeof(xl_stream) and whose read is set, and a non-NULL out_handle.";
                return false;
            }
            error = null;
            stream = new CallbackReadStream(*raw);
            return true;
        }

        public override int Read(Span<byte> buffer)
        {
            ObjectDisposedException.ThrowIf(Volatile.Read(ref _released) != 0, this);
            if (buffer.IsEmpty)
            {
                return 0;
            }
            if (!_peeked.IsEmpty)
            {
                int count = Math.Min(buffer.Length, _peeked.Length);
                _peeked.Span[..count].CopyTo(buffer);
                _peeked = _peeked[count..];
                return count;
            }
            SourceErrors.Clear();
            long got;
            fixed (byte* destination = buffer)
            {
                got = _read(_userData, destination, buffer.Length);
            }
            if (got < 0 || got > buffer.Length)
            {
                throw SourceErrors.Failed("xl_stream.read", got, buffer.Length);
            }
            return (int)got;
        }

        internal ReadOnlySpan<byte> Peek(int count)
        {
            byte[] head = new byte[count];
            int filled = 0;
            int read;
            while (filled < count && (read = Read(head.AsSpan(filled))) > 0)
            {
                filled += read;
            }
            _peeked = head.AsMemory(0, filled);
            return _peeked.Span;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return Read(buffer.AsSpan(offset, count));
        }

        protected override void Dispose(bool disposing)
        {
            if (Interlocked.Exchange(ref _released, 1) == 0 && _release is not null)
            {
                _release(_userData);
            }
            base.Dispose(disposing);
        }
    }
}
