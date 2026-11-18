using System.Buffers;

namespace ExcelReader.Core.Reader.Internal
{
    internal static class StreamSpooler
    {
        internal const int FirstChunk = 64 * 1024;
        internal const int MaxChunk = 1024 * 1024;

        internal static byte[] Spool(Stream stream, int firstChunk = FirstChunk, int maxChunk = MaxChunk)
        {
            List<(byte[] Buffer, int Length)> chunks = [];
            try
            {
                long total = 0;
                int size = firstChunk;
                while (true)
                {
                    byte[] buffer = ArrayPool<byte>.Shared.Rent(size);
                    int filled = 0;
                    int read = 1;
                    while (read > 0 && filled < size)
                    {
                        read = stream.Read(buffer, filled, size - filled);
                        filled += read;
                    }
                    chunks.Add((buffer, filled));
                    total += filled;
                    ThrowIfTooLong(total);
                    if (read == 0)
                    {
                        return Assemble(chunks, (int)total);
                    }
                    size = Math.Min(size * 2, maxChunk);
                }
            }
            finally
            {
                Release(chunks);
            }
        }

        internal static async ValueTask<byte[]> SpoolAsync(Stream stream, CancellationToken ct, int firstChunk = FirstChunk, int maxChunk = MaxChunk)
        {
            List<(byte[] Buffer, int Length)> chunks = [];
            try
            {
                long total = 0;
                int size = firstChunk;
                while (true)
                {
                    byte[] buffer = ArrayPool<byte>.Shared.Rent(size);
                    int filled = 0;
                    int read = 1;
                    while (read > 0 && filled < size)
                    {
                        read = await stream.ReadAsync(buffer.AsMemory(filled, size - filled), ct).ConfigureAwait(false);
                        filled += read;
                    }
                    chunks.Add((buffer, filled));
                    total += filled;
                    ThrowIfTooLong(total);
                    if (read == 0)
                    {
                        return Assemble(chunks, (int)total);
                    }
                    size = Math.Min(size * 2, maxChunk);
                }
            }
            finally
            {
                Release(chunks);
            }
        }

        private static void ThrowIfTooLong(long total)
        {
            if (total > Array.MaxLength)
            {
                throw new IOException("Stream was too long.");
            }
        }

        private static byte[] Assemble(List<(byte[] Buffer, int Length)> chunks, int total)
        {
            byte[] result = GC.AllocateUninitializedArray<byte>(total);
            int offset = 0;
            foreach ((byte[] buffer, int length) in chunks)
            {
                buffer.AsSpan(0, length).CopyTo(result.AsSpan(offset));
                offset += length;
            }
            return result;
        }

        private static void Release(List<(byte[] Buffer, int Length)> chunks)
        {
            foreach ((byte[] buffer, int _) in chunks)
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }
    }
}
