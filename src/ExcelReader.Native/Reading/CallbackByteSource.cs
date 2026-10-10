using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Reader.Sources;

namespace ExcelReader.Native.Reading
{
    internal sealed unsafe class CallbackByteSource : ByteSource
    {
        private readonly delegate* unmanaged<void*, long, byte*, long, long> _readAt;
        private readonly delegate* unmanaged<void*, void> _release;
        private readonly void* _userData;
        private int _released;

        private CallbackByteSource(in NativeSourceRaw raw)
        {
            _readAt = raw.ReadAt;
            _release = raw.Release;
            _userData = raw.UserData;
            Length = raw.Length;
        }

        internal override long Length { get; }

        internal static bool TryCreate(
            NativeSourceRaw* raw, bool hasOutHandle, [NotNullWhen(true)] out CallbackByteSource? source, [NotNullWhen(false)] out string? error)
        {
            source = null;
            if (raw is null || !hasOutHandle || raw->StructSize != sizeof(NativeSourceRaw) || raw->ReadAt is null || raw->Length < 0)
            {
                error = "xl_open_source needs a non-NULL source whose struct_size is sizeof(xl_source), whose read_at is set and whose length is 0 or more, and a non-NULL out_handle.";
                return false;
            }
            error = null;
            source = new CallbackByteSource(*raw);
            return true;
        }

        internal override int Read(long offset, Span<byte> destination)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(offset);
            ObjectDisposedException.ThrowIf(Volatile.Read(ref _released) != 0, this);
            if (offset >= Length || destination.IsEmpty)
            {
                return 0;
            }
            int requested = (int)Math.Min(destination.Length, Length - offset);
            SourceErrors.Clear();
            long got;
            fixed (byte* buffer = destination)
            {
                got = _readAt(_userData, offset, buffer, requested);
            }
            if (got < 0 || got > requested)
            {
                throw SourceErrors.Failed($"xl_source.read_at at offset {offset}", got, requested);
            }
            if (got == 0)
            {
                throw new IOException($"The source ended at offset {offset}, before its declared length of {Length} bytes.");
            }
            return (int)got;
        }

        internal override ValueTask<int> ReadAsync(long offset, Memory<byte> destination, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            return new ValueTask<int>(Read(offset, destination.Span));
        }

        public override void Dispose()
        {
            if (Interlocked.Exchange(ref _released, 1) == 0 && _release is not null)
            {
                _release(_userData);
            }
        }
    }
}
