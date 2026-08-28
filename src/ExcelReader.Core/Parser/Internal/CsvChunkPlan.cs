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
    // Chunk size is a small constant, not a share of the file. A chunk buffers its parsed models
    // until the ordered merge reaches it, so a chunk size that scales with the file scales peak heap
    // with it too. Keeping it small (64 KB, matching the record reader's own buffer) keeps in-flight
    // models small enough to die in Gen0 instead of being promoted.
    internal sealed class CsvChunkPlan
    {
        private const long DefaultChunkSize = 64 * 1024;

        private readonly long _dataStart;
        private readonly long _dataEnd;
        private readonly long _chunkSize;
        private readonly int _count;
        private int _next = -1;

        private CsvChunkPlan(long dataStart, long dataEnd, long chunkSize, int count)
        {
            _dataStart = dataStart;
            _dataEnd = dataEnd;
            _chunkSize = chunkSize;
            _count = count;
        }

        internal int Count => _count;

        // Computed rather than stored: chunks are uniform, and at 64 KB a 10 GB file has ~160k of them.
        internal CsvChunk this[int index]
        {
            get
            {
                long start = _dataStart + (index * _chunkSize);
                return new CsvChunk(index, start, Math.Min(start + _chunkSize, _dataEnd));
            }
        }

        // chunkSizeOverride exists so tests can force tiny chunks and walk a boundary across every
        // byte offset of a small fixture. Zero means "use the default".
        internal static CsvChunkPlan Create(long dataStart, long dataLength, int degreeOfParallelism, int chunkSizeOverride = 0)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(dataStart);
            ArgumentOutOfRangeException.ThrowIfNegative(dataLength);
            ArgumentOutOfRangeException.ThrowIfLessThan(degreeOfParallelism, 1);

            long chunkSize = chunkSizeOverride > 0 ? chunkSizeOverride : DefaultChunkSize;
            int count = (int)Math.Max(1, (dataLength + chunkSize - 1) / chunkSize);
            return new CsvChunkPlan(dataStart, dataStart + dataLength, chunkSize, count);
        }

        // Lock-free pull. Every worker calls this in a loop until it returns false.
        internal bool TryTakeNext(out CsvChunk chunk)
        {
            int index = Interlocked.Increment(ref _next);
            if (index >= _count)
            {
                chunk = default;
                return false;
            }
            chunk = this[index];
            return true;
        }
    }
}
