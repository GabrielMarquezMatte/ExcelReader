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
    // Chunk size is a small constant, NOT a share of the file. It was derived from the file length
    // (dataLength / (4 * dop)) until measurement showed that to be the parallel path's single
    // costliest decision: chunk size then grew without bound with the input, and since a chunk must
    // buffer its parsed models until the ordered merge reaches it, peak heap grew with it. On a
    // 200 MB corpus at dop=16 that was ~400 MB of live models and a stream of Gen2 collections; on a
    // 10 GB file it would have been gigabytes.
    //
    // Small chunks fix that at the root: in-flight models stay a few MB, so they die in Gen0 instead
    // of being promoted, and Gen2 collections disappear from the parallel path entirely. Measured on
    // a 16-core machine, 200 MB narrow-int / 140 MB conversion-heavy corpora, dop=8:
    //
    //   chunk size   narrow ms   wide ms   peak heap MB
    //   64 KB            254        294         15-19
    //   256 KB           257        338         17-21
    //   1 MB             335        621         33-50
    //   file/(4*dop)     569        704       272-321   (the original sizing)
    //
    // 64 KB also matches the record reader's own buffer, so a chunk is about one buffered read.
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

        // Chunks are uniform, so they are computed rather than stored. At 64 KB a 10 GB file has
        // ~160k of them, and an array of those would be megabytes of bookkeeping held live for the
        // whole enumeration for no information the index does not already carry.
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
