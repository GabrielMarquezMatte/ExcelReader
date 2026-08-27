using System.Runtime.InteropServices;

namespace ExcelReader.Core.Parser.Internal
{
    // One chunk's nominal byte bounds. Start is a *guess* for every chunk but index 0 — the true
    // first record start is confirmed by the predecessor during the ordered merge. End is where the
    // chunk stops accepting new records; the record straddling End is finished anyway (the worker
    // overshoots), and the next chunk skips it.
    [StructLayout(LayoutKind.Auto)]
    internal readonly record struct CsvChunk(int Index, long Start, long End);

    // A fixed split of the data range into chunks, plus the pull queue workers draw from.
    //
    // Chunks are deliberately NOT one-per-worker. With chunks equal to workers, emitting in order
    // means the last worker finishes early and its output sits buffered — keeping every worker busy
    // would require buffering up to (N-1)/N of the parsed file, which is untenable at 10 GB. At four
    // chunks per worker the reorder window is about one chunk per worker instead.
    internal sealed class CsvChunkPlan
    {
        private const int ChunksPerWorker = 4;
        private const int MinChunkSize = 64 * 1024;

        private readonly CsvChunk[] _chunks;
        private int _next = -1;

        private CsvChunkPlan(CsvChunk[] chunks)
        {
            _chunks = chunks;
        }

        internal int Count => _chunks.Length;

        internal CsvChunk this[int index] => _chunks[index];

        // chunkSizeOverride exists so tests can force tiny chunks and walk a boundary across every
        // byte offset of a small fixture. Zero means "derive from degreeOfParallelism".
        internal static CsvChunkPlan Create(long dataStart, long dataLength, int degreeOfParallelism, int chunkSizeOverride = 0)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(dataStart);
            ArgumentOutOfRangeException.ThrowIfNegative(dataLength);
            ArgumentOutOfRangeException.ThrowIfLessThan(degreeOfParallelism, 1);

            long chunkSize;
            if (chunkSizeOverride > 0)
            {
                chunkSize = chunkSizeOverride;
            }
            else
            {
                long target = dataLength / ((long)degreeOfParallelism * ChunksPerWorker);
                chunkSize = Math.Max(target, MinChunkSize);
            }

            int count = (int)Math.Max(1, (dataLength + chunkSize - 1) / chunkSize);
            var chunks = new CsvChunk[count];
            long cursor = dataStart;
            long end = dataStart + dataLength;
            for (int i = 0; i < count; i++)
            {
                long chunkEnd = i == count - 1 ? end : Math.Min(cursor + chunkSize, end);
                chunks[i] = new CsvChunk(i, cursor, chunkEnd);
                cursor = chunkEnd;
            }
            return new CsvChunkPlan(chunks);
        }

        // Lock-free pull. Every worker calls this in a loop until it returns false.
        internal bool TryTakeNext(out CsvChunk chunk)
        {
            int index = Interlocked.Increment(ref _next);
            if (index >= _chunks.Length)
            {
                chunk = default;
                return false;
            }
            chunk = _chunks[index];
            return true;
        }
    }
}
