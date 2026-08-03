using System.Runtime.CompilerServices;
using System.Text;

namespace ExcelReader.Core.ValueObjects
{
    /// <summary>
    /// Fixed-size, direct-mapped, content-keyed string cache: dedups repeated short text (categorical
    /// CSV/inline-string columns) without needing a stable table index the way the shared-string dedup
    /// cache does. A hash collision just evicts the older entry — the cached string is re-verified on
    /// every hit by re-encoding it back to UTF-8 into a stack buffer, so a collision can only cost cache
    /// effectiveness, never return the wrong string. No growth, no rehash, and — unlike keeping a
    /// separate copy of each key's raw bytes — a miss allocates nothing beyond the one <see cref="string"/>
    /// GetString() would have allocated anyway: high-cardinality (mostly-unique) columns pay no more than
    /// the no-cache baseline, so the cache is safe to leave on by default.
    /// </summary>
    internal sealed class Utf8StringCache
    {
        private const int BucketCount = 1024; // power of two, so hash & (BucketCount - 1) is a mask
        private const int MaxKeyLength = 64; // longer values rarely repeat; not worth caching
        private readonly string?[] _values = new string[BucketCount];

        /// <summary>Returns a cached, deduplicated string for <paramref name="utf8"/> when one is
        /// available; otherwise materializes and caches a new one (or, past <see cref="MaxKeyLength"/>,
        /// materializes without caching).</summary>
        [SkipLocalsInit]
        internal string GetOrAdd(ReadOnlySpan<byte> utf8)
        {
            if (utf8.IsEmpty)
            {
                return string.Empty;
            }
            if (utf8.Length > MaxKeyLength)
            {
                return Encoding.UTF8.GetString(utf8);
            }
            int bucket = ComputeHash(utf8) & (BucketCount - 1);
            string? cached = _values[bucket];
            if (cached is not null && Matches(cached, utf8))
            {
                return cached;
            }
            // Either no entry yet, or a bucket collision with a different value — either way,
            // overwrite unconditionally; a non-matching cached value must never be returned.
            string value = Encoding.UTF8.GetString(utf8);
            _values[bucket] = value;
            return value;
        }

        // Re-encodes the cached string back to UTF-8 into a stack buffer and compares bytes directly,
        // instead of keeping a persisted byte[] copy of the key — a round trip through a string can
        // never grow past its original UTF-8 byte length, so MaxKeyLength bounds this buffer too.
        [SkipLocalsInit]
        private static bool Matches(string cached, ReadOnlySpan<byte> utf8)
        {
            if (cached.Length > utf8.Length)
            {
                return false; // can't match: re-encoding could only add bytes, never fewer
            }
            Span<byte> buffer = stackalloc byte[MaxKeyLength];
            return Encoding.UTF8.TryGetBytes(cached, buffer, out int written)
                && buffer[..written].SequenceEqual(utf8);
        }

        private static int ComputeHash(ReadOnlySpan<byte> utf8)
        {
            HashCode hash = default;
            hash.AddBytes(utf8);
            return hash.ToHashCode();
        }
    }
}
