namespace ExcelReader.Core.Reader
{
    // Ref-struct enumerable over open XML tags whose names begin with a given prefix.
    // Replaces the IEnumerable<ReadOnlyMemory<byte>> Tags iterator so callers can work with
    // ReadOnlySpan<byte> throughout — no ReadOnlyMemory<byte> indirection, no heap allocation per tag.
    internal ref struct TagSpanEnumerable
    {
        private ReadOnlySpan<byte> _remaining;
        private readonly ReadOnlySpan<byte> _prefix;

        internal TagSpanEnumerable(ReadOnlySpan<byte> buf, ReadOnlySpan<byte> prefix)
        {
            _remaining = buf;
            _prefix = prefix;
        }

        // Pattern-based foreach: compiler calls GetEnumerator() once on the range expression,
        // then drives MoveNext() / Current on the returned copy.
        public readonly TagSpanEnumerable GetEnumerator()
        {
            return this;
        }

        public ReadOnlySpan<byte> Current { get; private set; }

        public bool MoveNext()
        {
            int start = _remaining.IndexOf(_prefix);
            if (start < 0)
            {
                return false;
            }
            int end = _remaining[start..].IndexOf((byte)'>');
            if (end < 0)
            {
                return false;
            }
            Current = _remaining.Slice(start, end + 1);
            _remaining = _remaining[(start + end + 1)..];
            return true;
        }
    }
}
