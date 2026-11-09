using System.Buffers.Binary;
using System.Runtime.CompilerServices;
using System.Text;

namespace ExcelReader.Core.Reader.Internal
{
    /// <summary>
    /// Fixed-size, 4-way set-associative, content-keyed string cache: dedups repeated short text (categorical
    /// CSV/inline-string columns) without needing a stable table index the way the shared-string dedup
    /// cache does. A full set evicts its oldest entry — the cached string is re-verified on
    /// every hit by re-encoding it back to UTF-8 into a stack buffer, so a collision can only cost cache
    /// effectiveness, never return the wrong string. No growth, no rehash, and a miss allocates nothing
    /// beyond the one <see cref="string"/> GetString() would have allocated anyway.
    /// </summary>
    /// <remarks>
    /// A direct-mapped cache let two hot values of a 10-value column evict each other and re-materialize a
    /// third of that column's strings; four ways keep any four values of a set cached. A lookup still costs
    /// about as much as materializing, so <see cref="ColumnStringCaches"/> gives each column its own instance
    /// and this pauses for <see cref="PausedCells"/> cells after a sample of lookups that rarely hit.
    /// </remarks>
    internal sealed class Utf8StringCache
    {
        private const int SetCount = 256;
        private const int Ways = 4;
        private const int MaxKeyLength = 64;
        private const int SampleLookups = 1024;
        private const int MinHitsPerSample = SampleLookups / 16;
        private const int PausedCells = 64 * 1024;
        private string?[]? _values;
        private int[]? _hashes;
        private int _lookups;
        private int _hits;
        private int _pausedCells;

        /// <summary>False while paused: the caller materializes without the cache. Counts down the pause.</summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal bool TryUse()
        {
            if (_pausedCells == 0)
            {
                return true;
            }
            _pausedCells--;
            return false;
        }

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
                return Utf8Strings.Create(utf8);
            }
            string?[] values = _values ??= new string[SetCount * Ways];
            int[] hashes = _hashes ??= new int[SetCount * Ways];
            int hash = ComputeHash(utf8);
            int start = (hash & (SetCount - 1)) * Ways;
            for (int i = start; i < start + Ways; i++)
            {
                string? cached = values[i];
                if (cached is null)
                {
                    break;
                }
                if (hashes[i] == hash && Matches(cached, utf8))
                {
                    Sample(hit: true);
                    return cached;
                }
            }
            Sample(hit: false);
            string value = Utf8Strings.Create(utf8);
            Array.Copy(values, start, values, start + 1, Ways - 1);
            Array.Copy(hashes, start, hashes, start + 1, Ways - 1);
            values[start] = value;
            hashes[start] = hash;
            return value;
        }

        private void Sample(bool hit)
        {
            _hits += hit ? 1 : 0;
            if (++_lookups < SampleLookups)
            {
                return;
            }
            _pausedCells = _hits < MinHitsPerSample ? PausedCells : 0;
            _lookups = 0;
            _hits = 0;
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

        // Keys are at most 64 bytes and every hit is re-verified, so a multiply-mix over 8-byte words is
        // enough; HashCode.AddBytes mixed 4 bytes per step and cost more than the string it saved.
        internal static int ComputeHash(ReadOnlySpan<byte> utf8)
        {
            const ulong Multiplier = 0x9E3779B97F4A7C15;
            ulong hash = (ulong)utf8.Length * Multiplier;
            while (utf8.Length >= sizeof(ulong))
            {
                hash = (hash ^ BinaryPrimitives.ReadUInt64LittleEndian(utf8)) * Multiplier;
                utf8 = utf8[sizeof(ulong)..];
            }
            ulong tail = 0;
            for (int i = 0; i < utf8.Length; i++)
            {
                tail |= (ulong)utf8[i] << (i * 8);
            }
            hash = (hash ^ tail) * Multiplier;
            return (int)(hash >> 32);
        }
    }
}
