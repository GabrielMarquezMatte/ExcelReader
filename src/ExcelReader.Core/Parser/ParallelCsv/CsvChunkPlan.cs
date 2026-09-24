using System.Runtime.InteropServices;

namespace ExcelReader.Core.Parser.ParallelCsv
{
    [StructLayout(LayoutKind.Auto)]
    internal readonly record struct CsvChunk(int Index, long Start, long End);

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

        internal CsvChunk this[int index]
        {
            get
            {
                long start = _dataStart + (index * _chunkSize);
                return new CsvChunk(index, start, Math.Min(start + _chunkSize, _dataEnd));
            }
        }

        internal static CsvChunkPlan Create(long dataStart, long dataLength, int degreeOfParallelism, int chunkSizeOverride = 0)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(dataStart);
            ArgumentOutOfRangeException.ThrowIfNegative(dataLength);
            ArgumentOutOfRangeException.ThrowIfLessThan(degreeOfParallelism, 1);

            long chunkSize = chunkSizeOverride > 0 ? chunkSizeOverride : DefaultChunkSize;
            int count = (int)Math.Max(1, (dataLength + chunkSize - 1) / chunkSize);
            return new CsvChunkPlan(dataStart, dataStart + dataLength, chunkSize, count);
        }

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
