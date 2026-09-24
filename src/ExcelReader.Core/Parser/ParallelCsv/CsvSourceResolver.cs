using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal static class CsvSourceResolver
    {
        internal static bool TryResolve(Stream stream, out CsvChunkSource source)
        {
            if (stream is MemoryStream memory && memory.TryGetBuffer(out ArraySegment<byte> segment))
            {
                int offset = segment.Offset + (int)memory.Position;
                int count = segment.Count - (int)memory.Position;
                if (count <= 0)
                {
                    source = default;
                    return false;
                }
                source = new CsvChunkSource(segment.Array!.AsMemory(offset, count));
                return true;
            }

            if (stream is FileStream file)
            {
                SafeFileHandle handle = file.SafeFileHandle;
                long length = RandomAccess.GetLength(handle);
                if (file.Position >= length)
                {
                    source = default;
                    return false;
                }
                source = new CsvChunkSource(handle, length, startOffset: file.Position);
                return true;
            }

            source = default;
            return false;
        }
    }
}
