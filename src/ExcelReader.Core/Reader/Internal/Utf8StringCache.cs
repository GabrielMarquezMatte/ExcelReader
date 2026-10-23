using System.Runtime.CompilerServices;
using System.Text;

namespace ExcelReader.Core.Reader.Internal
{
    /// <summary>
    /// Fixed-size, 4-way set-associative, content-keyed string cache: dedups repeated short text (categorical
    /// CSV/inline-string columns) without needing a stable table index the way the shared-string dedup
    /// cache does. A full set evicts its oldest entry — the cached string is re-verified on
    /// every hit by re-encoding it back to UTF-8 into a stack buffer, so a collision can only cost cache
    /// effectiveness, never return the wrong string. No growth, no rehash, and — unlike keeping a
    /// separate copy of each key's raw bytes — a miss allocates nothing beyond the one <see cref="string"/>
    /// GetString() would have allocated anyway: high-cardinality (mostly-unique) columns pay no more than
    /// the no-cache baseline, so the cache is safe to leave on by default.
    /// </summary>
    /// <remarks>
    /// <see cref="HashCode"/> is seeded per process, so which values share a set changes from run to run. A
    /// direct-mapped cache let two hot values of a 10-value column evict each other in ~3% of processes, and
    /// re-materialized a third of that column's strings there; four ways keep any four values of a set cached.
    /// </remarks>
    internal sealed class Utf8StringCache
    {
        private const int SetCount = 256;
        private const int Ways = 4;
        private const int MaxKeyLength = 64;
        private readonly string?[] _values = new string[SetCount * Ways];
        private readonly int[] _hashes = new int[SetCount * Ways];

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
            int hash = ComputeHash(utf8);
            int start = (hash & (SetCount - 1)) * Ways;
            for (int i = start; i < start + Ways; i++)
            {
                string? cached = _values[i];
                if (cached is null)
                {
                    break;
                }
                if (_hashes[i] == hash && Matches(cached, utf8))
                {
                    return cached;
                }
            }
            string value = Encoding.UTF8.GetString(utf8);
            Array.Copy(_values, start, _values, start + 1, Ways - 1);
            Array.Copy(_hashes, start, _hashes, start + 1, Ways - 1);
            _values[start] = value;
            _hashes[start] = hash;
            return value;
        }

        [SkipLocalsInit]
        private static bool Matches(string cached, ReadOnlySpan<byte> utf8)
        {
            if (cached.Length > utf8.Length)
            {
                return false;
            }
            if (cached.Length == utf8.Length && Ascii.Equals(cached, utf8))
            {
                return true;
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
