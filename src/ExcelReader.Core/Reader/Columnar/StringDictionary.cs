namespace ExcelReader.Core.Reader.Columnar
{
    /// <summary>
    /// The distinct strings of a dictionary column, in first-seen order, laid out as an Arrow utf8
    /// array. An open-addressing table over the stored bytes finds an existing entry.
    /// </summary>
    internal sealed class StringDictionary : IDisposable
    {
        private const int InitialSlots = 64;

        private readonly GrowableNativeBuffer<int> _offsets = NewOffsets();
        private readonly GrowableNativeBuffer<byte> _data = new();
        private readonly List<int> _hashes = [];
        private int[] _slots = new int[InitialSlots];

        internal int Count => _hashes.Count;

        internal GrowableNativeBuffer<int> Offsets => _offsets;

        internal GrowableNativeBuffer<byte> Data => _data;

        internal int GetOrAdd(ReadOnlySpan<byte> value)
        {
            int hash = Hash(value);
            int mask = _slots.Length - 1;
            int slot = hash & mask;
            while (_slots[slot] != 0)
            {
                int entry = _slots[slot] - 1;
                if (_hashes[entry] == hash && ValueAt(entry).SequenceEqual(value))
                {
                    return entry;
                }
                slot = (slot + 1) & mask;
            }
            int added = Count;
            _data.AddRange(value);
            _offsets.Add(_data.Count);
            _hashes.Add(hash);
            _slots[slot] = added + 1;
            if (Count * 2 > _slots.Length)
            {
                Rehash();
            }
            return added;
        }

        internal ReadOnlySpan<byte> ValueAt(int entry)
        {
            ReadOnlySpan<int> offsets = _offsets.AsSpan();
            return _data.AsSpan()[offsets[entry]..offsets[entry + 1]];
        }

        public void Dispose()
        {
            _offsets.Dispose();
            _data.Dispose();
        }

        private static GrowableNativeBuffer<int> NewOffsets()
        {
            GrowableNativeBuffer<int> offsets = new();
            offsets.Add(0);
            return offsets;
        }

        private static int Hash(ReadOnlySpan<byte> value)
        {
            HashCode hash = default;
            hash.AddBytes(value);
            return hash.ToHashCode();
        }

        private void Rehash()
        {
            int[] slots = new int[_slots.Length * 2];
            int mask = slots.Length - 1;
            for (int entry = 0; entry < _hashes.Count; entry++)
            {
                int slot = _hashes[entry] & mask;
                while (slots[slot] != 0)
                {
                    slot = (slot + 1) & mask;
                }
                slots[slot] = entry + 1;
            }
            _slots = slots;
        }
    }
}
