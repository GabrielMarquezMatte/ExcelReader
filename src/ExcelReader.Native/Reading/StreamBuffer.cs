using System.Buffers;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native.Reading
{
    internal static class StreamBuffer
    {
        internal static ReadOnlyMemory<byte> ReadAll(Stream stream, long limit)
        {
            limit = Math.Min(limit, Array.MaxLength);
            MemoryStream buffer = new();
            byte[] chunk = ArrayPool<byte>.Shared.Rent(81_920);
            try
            {
                int read;
                while ((read = stream.Read(chunk)) > 0)
                {
                    long total = buffer.Length + read;
                    if (total > limit)
                    {
                        throw new ExcelLimitExceededException("MaxBufferedBytes", limit, total);
                    }
                    buffer.Write(chunk, 0, read);
                }
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(chunk);
            }
            return buffer.GetBuffer().AsMemory(0, (int)buffer.Length);
        }
    }
}
