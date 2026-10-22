using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Parallel
{
    /// <summary>
    /// One worker's view of the workbook's shared-string table: texts it has already resolved are
    /// answered locally, and only a text new to this worker takes the table's lock.
    /// </summary>
    internal sealed class WorkerStringCache
    {
        private const int MaxEntries = 65_536;

        private readonly SharedStringTable _table;
        private readonly Lock _gate;
        private readonly Dictionary<string, int> _indexes = new(StringComparer.Ordinal);
        private readonly Dictionary<string, int>.AlternateLookup<ReadOnlySpan<char>> _spanIndexes;
        private int _localHits;

        internal WorkerStringCache(SharedStringTable table, Lock gate)
        {
            _table = table;
            _gate = gate;
            _spanIndexes = _indexes.GetAlternateLookup<ReadOnlySpan<char>>();
        }

        internal int GetOrAdd(string value)
        {
            if (_indexes.TryGetValue(value, out int index))
            {
                _localHits++;
                return index;
            }
            return Resolve(value);
        }

        internal int GetOrAdd(ReadOnlySpan<char> value)
        {
            if (_spanIndexes.TryGetValue(value, out int index))
            {
                _localHits++;
                return index;
            }
            lock (_gate)
            {
                index = _table.GetOrAdd(value);
            }
            if (_indexes.Count < MaxEntries)
            {
                _spanIndexes.TryAdd(value, index);
            }
            return index;
        }

        // The table counts every reference; ones answered locally are reported in bulk.
        internal void FlushReferences()
        {
            if (_localHits == 0)
            {
                return;
            }
            lock (_gate)
            {
                _table.AddReferences(_localHits);
            }
            _localHits = 0;
        }

        private int Resolve(ReadOnlySpan<char> value)
        {
            int index;
            lock (_gate)
            {
                index = _table.GetOrAdd(value);
            }
            if (_indexes.Count < MaxEntries)
            {
                _spanIndexes.TryAdd(value, index);
            }
            return index;
        }
    }
}
