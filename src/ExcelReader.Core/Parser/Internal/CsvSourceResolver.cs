using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Parser.Internal
{
    // Maps a caller's source onto one of the only two shapes a partition can be read from safely and
    // cheaply — a SafeFileHandle read positionally, or a byte range of memory — and declines
    // everything else.
    //
    // The declining is the important part. Accepting any CanSeek stream and serializing refills
    // behind a lock would read as more general, but CanSeek says nothing about what a seek COSTS:
    // this library's own DecryptedPackageStream is CanSeek with a known Length, yet every seek
    // re-decrypts AES segments. A narrow contract that says "these two types" is honest; a broad one
    // that silently performs terribly on the rest is not.
    internal static class CsvSourceResolver
    {
        // Returns false when the stream cannot be partitioned, in which case the caller parses
        // sequentially — same results, one thread.
        internal static bool TryResolve(Stream stream, out CsvChunkSource source)
        {
            if (stream is MemoryStream memory && memory.TryGetBuffer(out ArraySegment<byte> segment))
            {
                // Read from the stream's current position, matching Excel.FromCsv(Stream).
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

            // A MemoryStream whose buffer is not publicly visible falls through deliberately: the
            // only way to reach its bytes is ToArray(), a silent full copy of the source.
            if (stream is FileStream file)
            {
                // Reading SafeFileHandle flushes the FileStream's internal buffer and marks the
                // handle exposed, degrading that stream's later buffering. Documented on the public
                // overload. The handle is BORROWED — the caller owns the FileStream, so nothing here
                // ever disposes it.
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
