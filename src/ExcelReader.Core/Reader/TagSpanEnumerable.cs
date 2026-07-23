namespace ExcelReader.Core.Reader
{
    // Ref-struct enumerable over open XML tags whose names begin with a given prefix.
    // Replaces the IEnumerable<ReadOnlyMemory<byte>> Tags iterator so callers can work with
    // ReadOnlySpan<byte> throughout — no ReadOnlyMemory<byte> indirection, no heap allocation per tag.
    internal ref struct TagSpanEnumerable
    {
        private ReadOnlySpan<byte> _remaining;
        private readonly ReadOnlySpan<byte> _prefix;
        // When the caller's prefix doesn't already end at a name boundary (e.g. "<Relationship" with
        // no trailing space), a bare IndexOf would also match inside a longer element name like
        // "<RelationshipGroup". Prefixes that already end in a terminator (e.g. "<sheet ") are safe
        // by construction and skip the extra check.
        private readonly bool _needsBoundaryCheck;

        internal TagSpanEnumerable(ReadOnlySpan<byte> buf, ReadOnlySpan<byte> prefix)
        {
            _remaining = buf;
            _prefix = prefix;
            _needsBoundaryCheck = prefix.Length == 0 || !IsNameTerminator(prefix[^1]);
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
            while (true)
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
                if (!_needsBoundaryCheck)
                {
                    Current = _remaining.Slice(start, end + 1);
                    _remaining = _remaining[(start + end + 1)..];
                    return true;
                }
                int boundaryPos = start + _prefix.Length;
                bool atNameBoundary = boundaryPos >= _remaining.Length || IsNameTerminator(_remaining[boundaryPos]);
                if (!atNameBoundary)
                {
                    // Prefix matched inside a longer element name — not a real hit; keep scanning.
                    _remaining = _remaining[(start + 1)..];
                    continue;
                }
                Current = _remaining.Slice(start, end + 1);
                _remaining = _remaining[(start + end + 1)..];
                return true;
            }
        }

        private static bool IsNameTerminator(byte b)
        {
            return b is (byte)' ' or (byte)'>' or (byte)'/' or (byte)'\t' or (byte)'\r' or (byte)'\n';
        }
    }
}
