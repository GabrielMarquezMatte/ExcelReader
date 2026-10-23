using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Tests.Reader
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
            Assert.NotSame(cache.GetOrAdd(utf8), cache.GetOrAdd(utf8));
        }

        [Fact]
        public void ValuesSharingABucketStayCachedTogether()
        {
            byte[][] keys = KeysInOneBucket(2);
            var cache = new Utf8StringCache();
            string a = cache.GetOrAdd(keys[0]);
            string b = cache.GetOrAdd(keys[1]);
            Assert.Same(a, cache.GetOrAdd(keys[0]));
            Assert.Same(b, cache.GetOrAdd(keys[1]));
        }

        [Fact]
        public void OverfullBucketEvictsOldestEntryWithoutCorruptingAny()
        {
            byte[][] keys = KeysInOneBucket(5);
            var cache = new Utf8StringCache();
            string first = cache.GetOrAdd(keys[0]);
            foreach (byte[] key in keys[1..])
            {
                Assert.Equal(Encoding.UTF8.GetString(key), cache.GetOrAdd(key));
            }
            string again = cache.GetOrAdd(keys[0]);
            Assert.Equal(first, again);
            Assert.NotSame(first, again);
            foreach (byte[] key in keys)
            {
                Assert.Equal(Encoding.UTF8.GetString(key), cache.GetOrAdd(key));
            }
        }

        private static byte[][] KeysInOneBucket(int count)
        {
            var keys = new List<byte[]> { Encoding.UTF8.GetBytes("v0") };
            int bucket = BucketOf(keys[0]);
            for (int i = 1; keys.Count < count; i++)
            {
                byte[] candidate = Encoding.UTF8.GetBytes("v" + i.ToString(System.Globalization.CultureInfo.InvariantCulture));
                if (BucketOf(candidate) == bucket)
                {
                    keys.Add(candidate);
                }
            }
            return [.. keys];
        }

        private static int BucketOf(ReadOnlySpan<byte> utf8)
        {
            HashCode hash = default;
            hash.AddBytes(utf8);
            return hash.ToHashCode() & 1023;
        }

        [Fact]
        public void CsvInternStringsDefaultsToFalse()
        {
            using MemoryStream ms = new(Encoding.UTF8.GetBytes("alpha,1\nalpha,2\n"));
            using CsvReader reader = Excel.FromCsv(ms);
            using CsvReader.Enumerator e = reader.FirstSheet.GetEnumerator();

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
            using MemoryStream ms = new(Encoding.UTF8.GetBytes("alpha,1\nalpha,2\nbeta,3\n"));
            using CsvReader reader = Excel.FromCsv(ms, options: new ExcelReaderOptions { InternStrings = true });
            using CsvReader.Enumerator e = reader.FirstSheet.GetEnumerator();

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
            await using var reader = Excel.FromXlsx(ms);
            using var e = reader.FirstSheet.GetEnumerator();

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
            await using var reader = Excel.FromXlsx(ms, options: options);
            using var e = reader.FirstSheet.GetEnumerator();

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
