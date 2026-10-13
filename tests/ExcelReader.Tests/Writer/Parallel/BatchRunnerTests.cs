using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Parallel;

namespace ExcelReader.Tests.Writer.Parallel
{
    public class BatchRunnerTests
    {
        private sealed class Probe
        {
            private int _running;
            private int _peak;
            private int _formats;

            internal List<(int RowsBefore, int Rows)> Emitted { get; } = [];
            internal int Peak => _peak;
            internal int Running => Volatile.Read(ref _running);
            internal int Formats => Volatile.Read(ref _formats);
            internal int FailAtRowsBefore { get; init; } = -1;

            internal BatchSegment Format(int[] batch, int count, int rowsBefore)
            {
                int now = Interlocked.Increment(ref _running);
                Interlocked.Increment(ref _formats);
                InterlockedMax(ref _peak, now);
                try
                {
                    Thread.Sleep(2);
                    if (rowsBefore == FailAtRowsBefore)
                    {
                        throw new FormatException("batch " + rowsBefore);
                    }
                    Assert.Equal(rowsBefore, batch[0]);
                    return new BatchSegment(count) { RowFirst = rowsBefore };
                }
                finally
                {
                    Interlocked.Decrement(ref _running);
                }
            }

            internal ValueTask Emit(BatchSegment segment, CancellationToken ct)
            {
                Emitted.Add((segment.RowFirst, segment.Rows));
                return ValueTask.CompletedTask;
            }

            private static void InterlockedMax(ref int target, int value)
            {
                int seen = Volatile.Read(ref target);
                while (value > seen)
                {
                    int previous = Interlocked.CompareExchange(ref target, value, seen);
                    if (previous == seen)
                    {
                        return;
                    }
                    seen = previous;
                }
            }
        }

        private static BatchRunner<int> Runner(Probe probe, int degree, int batchRows, CancellationToken ct, int rowsBefore = 0)
        {
            var options = new ParallelWriteOptions { DegreeOfParallelism = degree, BatchRows = batchRows };
            return new BatchRunner<int>(options, rowsBefore, probe.Format, probe.Emit, ct);
        }

        private static async IAsyncEnumerable<int> Async(int count)
        {
            for (int i = 0; i < count; i++)
            {
                await Task.Yield();
                yield return i;
            }
        }

        [Theory]
        [InlineData(0, 0)]
        [InlineData(1, 1)]
        [InlineData(10, 1)]
        [InlineData(11, 2)]
        [InlineData(95, 10)]
        public async Task BatchesComeOutInOrderWithAPartialLastOne(int records, int expectedBatches)
        {
            foreach (int degree in new[] { 1, 4 })
            {
                var probe = new Probe();
                await Runner(probe, degree, 10, TestContext.Current.CancellationToken)
                    .RunAsync(new RecordSource<int>(Enumerable.Range(0, records)));

                Assert.Equal(expectedBatches, probe.Emitted.Count);
                Assert.Equal(records, probe.Emitted.Sum(static e => e.Rows));
                Assert.Equal(probe.Emitted.OrderBy(static e => e.RowsBefore), probe.Emitted);
                Assert.True(probe.Peak <= degree);
            }
        }

        [Fact]
        public async Task AnAsyncSourceGivesTheSameBatches()
        {
            var probe = new Probe();
            await Runner(probe, 3, 10, TestContext.Current.CancellationToken).RunAsync(new RecordSource<int>(Async(45)));

            Assert.Equal([(0, 10), (10, 10), (20, 10), (30, 10), (40, 5)], probe.Emitted);
        }

        [Fact]
        public async Task RowsBeforeOffsetsEveryBatch()
        {
            var probe = new Probe();
            int[] source = [.. Enumerable.Range(7, 25)];
            await Runner(probe, 2, 10, TestContext.Current.CancellationToken, rowsBefore: 7).RunAsync(new RecordSource<int>(source));

            Assert.Equal([(7, 10), (17, 10), (27, 5)], probe.Emitted);
        }

        [Fact]
        public async Task AFailingBatchStopsTheRunAndLeavesNoWorker()
        {
            var probe = new Probe { FailAtRowsBefore = 30 };
            BatchRunner<int> runner = Runner(probe, 4, 10, TestContext.Current.CancellationToken);

            FormatException failure = await Assert.ThrowsAsync<FormatException>(
                async () => await runner.RunAsync(new RecordSource<int>(Enumerable.Range(0, 500))));

            Assert.Equal("batch 30", failure.Message);
            Assert.Equal(0, probe.Running);
            Assert.All(probe.Emitted, static e => Assert.True(e.RowsBefore < 30));
        }

        [Fact]
        public async Task AThrowingSourceSurfacesItsOwnException()
        {
            static IEnumerable<int> Source()
            {
                for (int i = 0; i < 25; i++)
                {
                    yield return i;
                }
                throw new InvalidDataException("source");
            }
            var probe = new Probe();

            await Assert.ThrowsAsync<InvalidDataException>(
                async () => await Runner(probe, 4, 10, TestContext.Current.CancellationToken).RunAsync(new RecordSource<int>(Source())));
            Assert.Equal(0, probe.Running);
        }

        [Fact]
        public async Task CancellationStopsBetweenBatches()
        {
            using var cts = new CancellationTokenSource();
            var probe = new Probe();
            IEnumerable<int> Source()
            {
                for (int i = 0; i < 1_000; i++)
                {
                    if (i == 35)
                    {
                        cts.Cancel();
                    }
                    yield return i;
                }
            }

            await Assert.ThrowsAnyAsync<OperationCanceledException>(
                async () => await Runner(probe, 2, 10, cts.Token).RunAsync(new RecordSource<int>(Source())));
            Assert.Equal(0, probe.Running);
            Assert.True(probe.Emitted.Count < 100);
        }

        [Fact]
        public async Task TheRowLimitThrowsBeforeTheOverflowingBatchIsFormatted()
        {
            var probe = new Probe();
            int start = 1_048_576 - 15;
            int[] source = [.. Enumerable.Range(start, 30)];

            await Assert.ThrowsAsync<ExcelLimitExceededException>(
                async () => await Runner(probe, 1, 10, TestContext.Current.CancellationToken, rowsBefore: start).RunAsync(new RecordSource<int>(source)));
            Assert.Equal(1, probe.Formats);
            Assert.DoesNotContain(probe.Emitted, e => e.RowsBefore > start);
        }

        [Fact]
        public void OptionsRejectValuesOutOfRange()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new ParallelWriteOptions { DegreeOfParallelism = -1 }.Validate());
            Assert.Throws<ArgumentOutOfRangeException>(() => new ParallelWriteOptions { BatchRows = 0 }.Validate());
            ParallelWriteOptions.Default.Validate();
            Assert.Equal(8192, ParallelWriteOptions.Default.BatchRows);
            Assert.Equal(0, ParallelWriteOptions.Default.DegreeOfParallelism);
        }
    }
}
