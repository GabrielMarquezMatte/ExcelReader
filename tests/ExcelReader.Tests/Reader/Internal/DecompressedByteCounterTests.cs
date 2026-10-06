using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Zip;

namespace ExcelReader.Tests.Reader.Internal
{
    public sealed class DecompressedByteCounterTests
    {
        [Fact]
        public void Concurrent_Adds_Are_All_Counted()
        {
            const int Threads = 16;
            const int PerThread = 20_000;
            DecompressedByteCounter counter = new(Threads * PerThread);

            Parallel.For(0, Threads, new ParallelOptions { MaxDegreeOfParallelism = Threads }, _ =>
            {
                for (int i = 0; i < PerThread; i++)
                {
                    counter.Add(1);
                }
            });

            Assert.Equal(0, counter.Remaining);
            Assert.Throws<ExcelLimitExceededException>(() => counter.Add(1));
        }

        [Fact]
        public void Passing_The_Limit_On_One_Thread_Fails_The_Others_Too()
        {
            DecompressedByteCounter counter = new(100);
            Assert.Throws<ExcelLimitExceededException>(() => counter.Add(101));
            Assert.Throws<ExcelLimitExceededException>(() => counter.Add(1));
        }

        [Fact]
        public void A_Counter_Without_A_Limit_Never_Throws()
        {
            DecompressedByteCounter counter = new(0);
            counter.Add(long.MaxValue);
            counter.Add(long.MaxValue);
            Assert.Equal(long.MaxValue, counter.Remaining);
        }
    }
}
