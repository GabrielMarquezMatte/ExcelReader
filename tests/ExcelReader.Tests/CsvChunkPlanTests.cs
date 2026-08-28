using ExcelReader.Core.Parser.Internal;

namespace ExcelReader.Tests
{
    public class CsvChunkPlanTests
    {
        // Chunk size is a constant, not a share of the file: doubling the input doubles the chunk
        // count and leaves chunk size — and therefore the parallel path's peak memory — untouched.
        [Theory]
        [InlineData(16 * 1024 * 1024, 256)]
        [InlineData(32 * 1024 * 1024, 512)]
        public void SplitsIntoFixedSizeChunksIndependentOfFileSize(int dataLength, int expectedChunks)
        {
            CsvChunkPlan plan = CsvChunkPlan.Create(dataStart: 0, dataLength: dataLength, degreeOfParallelism: 4);

            Assert.Equal(expectedChunks, plan.Count);
            Assert.Equal(64 * 1024, plan[0].End - plan[0].Start);
        }

        [Fact]
        public void ChunkSizeDoesNotChangeWithDegreeOfParallelism()
        {
            CsvChunkPlan one = CsvChunkPlan.Create(dataStart: 0, dataLength: 16 * 1024 * 1024, degreeOfParallelism: 1);
            CsvChunkPlan many = CsvChunkPlan.Create(dataStart: 0, dataLength: 16 * 1024 * 1024, degreeOfParallelism: 32);

            Assert.Equal(one.Count, many.Count);
        }

        [Fact]
        public void ChunksCoverTheWholeRangeContiguouslyWithNoGapsOrOverlap()
        {
            CsvChunkPlan plan = CsvChunkPlan.Create(dataStart: 100, dataLength: 1000, degreeOfParallelism: 3);

            Assert.Equal(100, plan[0].Start);
            for (int i = 1; i < plan.Count; i++)
            {
                Assert.Equal(plan[i - 1].End, plan[i].Start);
            }
            Assert.Equal(1100, plan[plan.Count - 1].End);
        }

        [Fact]
        public void ChunkIndicesAreSequentialFromZero()
        {
            CsvChunkPlan plan = CsvChunkPlan.Create(dataStart: 0, dataLength: 999, degreeOfParallelism: 2);

            for (int i = 0; i < plan.Count; i++)
            {
                Assert.Equal(i, plan[i].Index);
            }
        }

        [Fact]
        public void AnOverrideForcesTinyChunksForBoundaryTesting()
        {
            CsvChunkPlan plan = CsvChunkPlan.Create(dataStart: 0, dataLength: 256, degreeOfParallelism: 2, chunkSizeOverride: 64);

            Assert.Equal(4, plan.Count);
            Assert.Equal(0, plan[0].Start);
            Assert.Equal(64, plan[0].End);
            Assert.Equal(256, plan[3].End);
        }

        [Fact]
        public void ARangeSmallerThanOneChunkBecomesASingleChunk()
        {
            CsvChunkPlan plan = CsvChunkPlan.Create(dataStart: 0, dataLength: 10, degreeOfParallelism: 8);

            Assert.Equal(1, plan.Count);
            Assert.Equal(0, plan[0].Start);
            Assert.Equal(10, plan[0].End);
        }

        [Fact]
        public void EveryChunkIsHandedOutExactlyOnceUnderConcurrentPulls()
        {
            CsvChunkPlan plan = CsvChunkPlan.Create(dataStart: 0, dataLength: 100_000, degreeOfParallelism: 8);
            int[] takenCount = new int[plan.Count];

            Parallel.For(0, 16, _ =>
            {
                while (plan.TryTakeNext(out CsvChunk chunk))
                {
                    Interlocked.Increment(ref takenCount[chunk.Index]);
                }
            });

            Assert.All(takenCount, c => Assert.Equal(1, c));
        }
    }
}
