using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Reader.Sources
{
    /// <summary>
    /// Random-access bytes of one workbook file. Every read names its own offset, so readers of the
    /// same source share no cursor.
    /// </summary>
    internal abstract class ByteSource : IDisposable
    {
        internal abstract long Length { get; }

        internal abstract int Read(long offset, Span<byte> destination);

        internal abstract ValueTask<int> ReadAsync(long offset, Memory<byte> destination, CancellationToken ct);

        internal virtual bool TryGetMemory(out ReadOnlyMemory<byte> memory)
        {
            memory = default;
            return false;
        }

        internal void ReadExactly(long offset, Span<byte> destination)
        {
            int filled = 0;
            while (filled < destination.Length)
            {
                int read = Read(offset + filled, destination[filled..]);
                if (read <= 0)
                {
                    throw new EndOfStreamException();
                }
                filled += read;
            }
        }

        internal async ValueTask ReadExactlyAsync(long offset, Memory<byte> destination, CancellationToken ct)
        {
            int filled = 0;
            while (filled < destination.Length)
            {
                int read = await ReadAsync(offset + filled, destination[filled..], ct).ConfigureAwait(false);
                if (read <= 0)
                {
                    throw new EndOfStreamException();
                }
                filled += read;
            }
        }

        public abstract void Dispose();

        internal static ByteSource FromMemory(ReadOnlyMemory<byte> data)
        {
            return new MemoryByteSource(data, owner: null);
        }

        internal static ByteSource OpenFile(string path, bool asynchronous)
        {
            SafeFileHandle handle = File.OpenHandle(path, FileMode.Open, FileAccess.Read, FileShare.Read,
                asynchronous ? FileOptions.Asynchronous : FileOptions.None);
            try
            {
                return new FileByteSource(handle, RandomAccess.GetLength(handle), owner: handle);
            }
            catch
            {
                handle.Dispose();
                throw;
            }
        }

        internal static ByteSource FromStream(Stream stream, bool leaveOpen)
        {
            if (TryWrapSeekable(stream, leaveOpen, out ByteSource? source))
            {
                return source;
            }
            try
            {
                using MemoryStream copy = new();
                stream.CopyTo(copy);
                return new MemoryByteSource(copy.GetBuffer().AsMemory(0, (int)copy.Length), owner: null);
            }
            finally
            {
                if (!leaveOpen)
                {
                    stream.Dispose();
                }
            }
        }

        internal static async ValueTask<ByteSource> FromStreamAsync(Stream stream, bool leaveOpen, CancellationToken ct)
        {
            if (TryWrapSeekable(stream, leaveOpen, out ByteSource? source))
            {
                return source;
            }
            try
            {
                using MemoryStream copy = new();
                await stream.CopyToAsync(copy, ct).ConfigureAwait(false);
                return new MemoryByteSource(copy.GetBuffer().AsMemory(0, (int)copy.Length), owner: null);
            }
            finally
            {
                if (!leaveOpen)
                {
                    await stream.DisposeAsync().ConfigureAwait(false);
                }
            }
        }

        private static bool TryWrapSeekable(Stream stream, bool leaveOpen, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out ByteSource? source)
        {
            IDisposable? owner = leaveOpen ? null : stream;
            if (stream is MemoryStream memory && memory.TryGetBuffer(out ArraySegment<byte> segment))
            {
                source = new MemoryByteSource(segment, owner);
                return true;
            }
            if (stream is FileStream file && file.CanSeek)
            {
                source = new FileByteSource(file.SafeFileHandle, file.Length, owner);
                return true;
            }
            if (stream.CanSeek)
            {
                source = new StreamByteSource(stream, leaveOpen);
                return true;
            }
            source = null;
            return false;
        }
    }
}
