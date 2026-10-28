using System.Globalization;
using ExcelReader.Core.Writer.Internal;
using ExcelReader.Core.Writer.Parallel;

namespace ExcelReader.Tests.Writer.Parallel
{
    public class WorkerStringCacheTests
    {
        [Fact]
        public void GetOrAdd_Should_Not_Allocate_For_A_Text_The_Table_Already_Holds_When_The_Local_Cache_Is_Full()
        {
            var table = new SharedStringTable();
            var cache = new WorkerStringCache(table);
            for (int i = 0; i < 65_536; i++)
            {
                cache.GetOrAdd(i.ToString(CultureInfo.InvariantCulture).AsSpan());
            }
            int expected = table.GetOrAdd("already-shared");
            ReadOnlySpan<char> text = "already-shared".AsSpan();

            cache.GetOrAdd(text);
            long before = GC.GetAllocatedBytesForCurrentThread();
            int index = cache.GetOrAdd(text);
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

            Assert.Equal(expected, index);
            Assert.Equal(0, allocated);
            Assert.Equal(65_537, table.UniqueCount);
        }

        [Fact]
        public void GetOrAdd_Should_Answer_A_Span_From_The_Local_Cache_And_Count_References_On_Flush()
        {
            var table = new SharedStringTable();
            var cache = new WorkerStringCache(table);

            int first = cache.GetOrAdd("widget".AsSpan());
            int second = cache.GetOrAdd("widget");

            Assert.Equal(first, second);
            Assert.Equal(1, table.UniqueCount);
            Assert.Equal(0, table.Count);
            cache.FlushReferences();
            Assert.Equal(2, table.Count);
        }

        [Fact]
        public void A_Text_Added_By_One_Path_Should_Keep_Its_Index_On_The_Other()
        {
            var table = new SharedStringTable();

            int serialA = table.GetOrAdd("a");
            int concurrentA = table.GetOrAddConcurrent("a", "a");
            int concurrentB = table.GetOrAddConcurrent("b", owned: null);
            int serialB = table.GetOrAdd("b".AsSpan());
            int serialC = table.GetOrAdd("c");

            Assert.Equal([0, 0, 1, 1, 2], [serialA, concurrentA, concurrentB, serialB, serialC]);
            Assert.Equal(3, table.UniqueCount);
            using MemoryStream xml = new();
            table.WriteTo(xml, binary: false);
            Assert.Contains("<si><t>a</t></si><si><t>b</t></si><si><t>c</t></si>", System.Text.Encoding.UTF8.GetString(xml.ToArray()), StringComparison.Ordinal);
        }

        [Fact]
        public void GetOrAddConcurrent_Should_Give_Every_Thread_The_Same_Dense_Indexes()
        {
            const int texts = 20_000;
            const int threads = 8;
            var table = new SharedStringTable();
            string[] values = [.. Enumerable.Range(0, texts).Select(i => "text-" + i.ToString(CultureInfo.InvariantCulture))];
            int[][] seen = new int[threads][];

            System.Threading.Tasks.Parallel.For(0, threads, new ParallelOptions { MaxDegreeOfParallelism = threads }, thread =>
            {
                int[] indexes = new int[texts];
                for (int step = 0; step < texts; step++)
                {
                    int i = (step * 7919 + thread * 104_729) % texts;
                    indexes[i] = table.GetOrAddConcurrent(values[i], thread % 2 == 0 ? values[i] : null);
                }
                seen[thread] = indexes;
            });

            Assert.Equal(texts, table.UniqueCount);
            Assert.All(seen, indexes => Assert.Equal(seen[0], indexes));
            Assert.Equal(Enumerable.Range(0, texts), seen[0].Order());
        }
    }
}
