using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Parallel
{
    /// <summary>
    /// One worker's view of the workbook's shared-string table: texts it has already resolved are
    /// answered locally, and only a text new to this worker takes one of the table's shard locks.
    /// </summary>
    internal sealed class WorkerStringCache
    {
        private const int MaxEntries = 65_536;

        private readonly SharedStringTable _table;
        private readonly Dictionary<string, int> _indexes = new(StringComparer.Ordinal);
        private readonly Dictionary<string, int>.AlternateLookup<ReadOnlySpan<char>> _spanIndexes;
        private int _references;

        internal WorkerStringCache(SharedStringTable table)
        {
            _table = table;
            _spanIndexes = _indexes.GetAlternateLookup<ReadOnlySpan<char>>();
        }

        internal int GetOrAdd(string value)
        {
            return GetOrAdd(value.AsSpan());
        }

        internal int GetOrAdd(ReadOnlySpan<char> value)
        {
            _references++;
            if (_spanIndexes.TryGetValue(value, out int index))
            {
                return index;
            }
            index = _table.GetOrAddConcurrent(value);
            if (_indexes.Count < MaxEntries)
            {
                _spanIndexes.TryAdd(value, index);
            }
            return index;
        }

        // The table counts every reference; this worker's are reported in bulk.
        internal void FlushReferences()
        {
            if (_references == 0)
            {
                return;
            }
            _table.AddReferences(_references);
            _references = 0;
        }
    }
}
