namespace ExcelReader.Core.Reader.Xlsx
{
    internal ref struct TagSpanEnumerable
    {
        private ReadOnlySpan<byte> _remaining;
        private readonly ReadOnlySpan<byte> _prefix;
        private readonly bool _needsBoundaryCheck;

        internal TagSpanEnumerable(ReadOnlySpan<byte> buf, ReadOnlySpan<byte> prefix)
        {
            _remaining = buf;
            _prefix = prefix;
            _needsBoundaryCheck = prefix.Length == 0 || !IsNameTerminator(prefix[^1]);
        }

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
