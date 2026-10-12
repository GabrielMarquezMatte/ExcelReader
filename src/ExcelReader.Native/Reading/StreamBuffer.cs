using ExcelReader.Core.Reader;

namespace ExcelReader.Native.Reading
{
    internal static class StreamBuffer
    {
        private const int FirstSegmentBytes = 1024 * 1024;
        private const int MaxSegmentBytes = 64 * 1024 * 1024;

        internal static ReadOnlyMemory<byte> ReadAll(Stream stream, long limit)
        {
            limit = Math.Min(limit, Array.MaxLength);
            List<byte[]> filledSegments = [];
            byte[] segment = GC.AllocateUninitializedArray<byte>(FirstSegmentBytes);
            int used = 0;
            long total = 0;
            int read;
            while ((read = stream.Read(segment.AsSpan(used))) > 0)
            {
                used += read;
                total += read;
                if (total > limit)
                {
                    throw new ExcelLimitExceededException("MaxBufferedBytes", limit, total);
                }
                if (used == segment.Length)
                {
                    filledSegments.Add(segment);
                    segment = GC.AllocateUninitializedArray<byte>(Math.Min(segment.Length * 2, MaxSegmentBytes));
                    used = 0;
                }
            }
            byte[] data = GC.AllocateUninitializedArray<byte>((int)total);
            int offset = 0;
            foreach (byte[] filled in filledSegments)
            {
                filled.CopyTo(data, offset);
                offset += filled.Length;
            }
            segment.AsSpan(0, used).CopyTo(data.AsSpan(offset));
            return data;
        }
    }
}
