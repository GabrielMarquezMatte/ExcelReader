using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Tests
{
    public class Utf8StringCacheTests
    {
        [Fact]
        public void RepeatedValueReturnsSameInstance()
        {
            var cache = new Utf8StringCache();
            byte[] utf8 = Encoding.UTF8.GetBytes("alpha");
            string first = cache.GetOrAdd(utf8);
            string second = cache.GetOrAdd(utf8);
            Assert.Same(first, second);
        }

        [Fact]
        public void DistinctValuesReturnCorrectText()
        {
            var cache = new Utf8StringCache();
            Assert.Equal("alpha", cache.GetOrAdd(Encoding.UTF8.GetBytes("alpha")));
            Assert.Equal("beta", cache.GetOrAdd(Encoding.UTF8.GetBytes("beta")));
            Assert.Equal("alpha", cache.GetOrAdd(Encoding.UTF8.GetBytes("alpha")));
        }

        [Fact]
        public void EmptyValueReturnsEmptyString()
        {
            var cache = new Utf8StringCache();
            Assert.Equal(string.Empty, cache.GetOrAdd(ReadOnlySpan<byte>.Empty));
        }

        [Fact]
        public void ValueLongerThanCapBypassesCacheButStillCorrect()
        {
            var cache = new Utf8StringCache();
            string longValue = new string('x', 200);
            byte[] utf8 = Encoding.UTF8.GetBytes(longValue);
            Assert.Equal(longValue, cache.GetOrAdd(utf8));
            // Not the same instance twice, since values past the cap are never cached — this is the
            // documented tradeoff, not a bug.
            Assert.NotSame(cache.GetOrAdd(utf8), cache.GetOrAdd(utf8));
        }

        [Fact]
        public void BucketCollisionEvictsOlderEntryWithoutCorruptingEither()
        {
            // Two distinct values landing in the same bucket must never return each other's text —
            // GetOrAdd re-verifies the cached bytes on every hit, so a collision only costs cache
            // effectiveness, never correctness. Mirrors the cache's own HashCode.AddBytes/1024-bucket
            // scheme to search for a real collision instead of assuming internal layout.
            const int bucketCount = 1024;
            byte[] keyA = Encoding.UTF8.GetBytes("v0");
            int bucketOfA = BucketOf(keyA, bucketCount);
            byte[] keyB = keyA;
            int i = 1;
            while (true)
            {
                byte[] candidate = Encoding.UTF8.GetBytes("v" + i.ToString(System.Globalization.CultureInfo.InvariantCulture));
                if (BucketOf(candidate, bucketCount) == bucketOfA && !candidate.AsSpan().SequenceEqual(keyA))
                {
                    keyB = candidate;
                    break;
                }
                i++;
            }
            var cache = new Utf8StringCache();
            string a = Encoding.UTF8.GetString(keyA);
            string b = Encoding.UTF8.GetString(keyB);
            Assert.Equal(a, cache.GetOrAdd(keyA));
            Assert.Equal(b, cache.GetOrAdd(keyB));
            Assert.Equal(a, cache.GetOrAdd(keyA));
        }

        private static int BucketOf(ReadOnlySpan<byte> utf8, int bucketCount)
        {
            HashCode hash = default;
            hash.AddBytes(utf8);
            return hash.ToHashCode() & (bucketCount - 1);
        }

        [Fact]
        public void CsvInternStringsDefaultsToFalse()
        {
            // Off by default (see CsvReaderOptions.InternStrings) — a real-world, mixed-cardinality
            // corpus measured this as a net loss when left on unconditionally, so a caller who hasn't
            // explicitly opted in must never see cache-driven reference-equality behavior.
            using MemoryStream ms = new(Encoding.UTF8.GetBytes("alpha,1\nalpha,2\n"));
            using CsvReader reader = Excel.FromCsv(ms);
            using CsvReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            string row1 = e.Current[0].GetString();
            Assert.True(e.MoveNext());
            string row2 = e.Current[0].GetString();

            Assert.Equal(row1, row2);
            Assert.NotSame(row1, row2);
        }

        [Fact]
        public void CsvInternStringsTrueEnablesDedupAcrossRows()
        {
            var options = new CsvReaderOptions { InternStrings = true };
            using MemoryStream ms = new(Encoding.UTF8.GetBytes("alpha,1\nalpha,2\nbeta,3\n"));
            using CsvReader reader = Excel.FromCsv(ms, options: options);
            using CsvReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            string row1 = e.Current[0].GetString();
            Assert.True(e.MoveNext());
            string row2 = e.Current[0].GetString();
            Assert.True(e.MoveNext());
            string row3 = e.Current[0].GetString();

            Assert.Same(row1, row2);
            Assert.Equal("alpha", row1);
            Assert.Equal("beta", row3);
            Assert.NotSame(row1, row3);
        }

        [Fact]
        public async Task XlsxInternStringsDefaultsToFalse()
        {
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["alpha"], ["alpha"], ["beta"]);
            await using var reader = Excel.From(ms);
            using var e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            string row1 = e.Current[0].GetString();
            Assert.True(e.MoveNext());
            string row2 = e.Current[0].GetString();

            Assert.Equal(row1, row2);
            Assert.NotSame(row1, row2);
        }

        [Fact]
        public async Task XlsxInternStringsTrueEnablesDedupForInlineStrings()
        {
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["alpha"], ["alpha"], ["beta"]);
            var options = new ExcelReaderOptions { InternStrings = true };
            await using var reader = Excel.From(ms, options: options);
            using var e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            string row1 = e.Current[0].GetString();
            Assert.True(e.MoveNext());
            string row2 = e.Current[0].GetString();
            Assert.True(e.MoveNext());
            string row3 = e.Current[0].GetString();

            Assert.Same(row1, row2);
            Assert.Equal("alpha", row1);
            Assert.Equal("beta", row3);
        }
    }
}
