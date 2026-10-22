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
            var cache = new WorkerStringCache(table, new Lock());
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
        public void GetOrAdd_Should_Answer_A_Span_From_The_Local_Cache_After_The_First_Lookup()
        {
            var table = new SharedStringTable();
            var cache = new WorkerStringCache(table, new Lock());

            int first = cache.GetOrAdd("widget".AsSpan());
            int second = cache.GetOrAdd("widget".AsSpan());

            Assert.Equal(first, second);
            Assert.Equal(1, table.UniqueCount);
            Assert.Equal(1, table.Count);
            cache.FlushReferences();
            Assert.Equal(2, table.Count);
        }
    }
}
