using System.Collections.Concurrent;
using System.Diagnostics.Tracing;
using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser.ParallelCsv;
using ExcelReader.Core.Parser.SheetPipeline;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Parser.SheetPipeline
{
    [CollectionDefinition(nameof(ArrayPoolEvents), DisableParallelization = true)]
    public sealed class ArrayPoolEvents
    {
    }

    // ArrayPool<byte>.Shared reports every rent and return through its event source; running alone keeps
    // other tests' rentals of the same size out of the counts.
    [Collection(nameof(ArrayPoolEvents))]
    public sealed class ChunkOwnershipTests
    {
        // Rounds up to a 512 KiB bucket, which nothing else in these tests rents.
        private const int ChunkBytes = 300_000;
        private const int PooledChunkSize = 512 * 1024;

        private sealed class PoolEvents : EventListener
        {
            internal readonly ConcurrentDictionary<int, int> Outstanding = new();
            internal readonly ConcurrentBag<int> Returned = [];

            protected override void OnEventSourceCreated(EventSource eventSource)
            {
                if (string.Equals(eventSource.Name, "System.Buffers.ArrayPoolEventSource", StringComparison.Ordinal))
                {
                    EnableEvents(eventSource, EventLevel.Verbose);
                }
            }

            protected override void OnEventWritten(EventWrittenEventArgs eventData)
            {
                if (eventData.Payload is not { Count: >= 2 } payload || payload[0] is not int id || payload[1] is not int size)
                {
                    return;
                }
                if (string.Equals(eventData.EventName, "BufferReturned", StringComparison.Ordinal))
                {
                    Returned.Add(id);
                }
                if (size != PooledChunkSize)
                {
                    return;
                }
                if (string.Equals(eventData.EventName, "BufferRented", StringComparison.Ordinal))
                {
                    Outstanding.AddOrUpdate(id, 1, static (_, count) => count + 1);
                }
                if (string.Equals(eventData.EventName, "BufferReturned", StringComparison.Ordinal))
                {
                    Outstanding.AddOrUpdate(id, -1, static (_, count) => count - 1);
                }
            }
        }

        private static byte[] SheetXml(int rows)
        {
            var sb = new StringBuilder("<worksheet><sheetData>");
            for (int r = 1; r <= rows; r++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"<row r=\"{r}\"><c><v>{r}</v></c></row>");
            }
            return Encoding.UTF8.GetBytes(sb.Append("</sheetData></worksheet>").ToString());
        }

        private static async Task<int> RejectedEmitAsync(byte[] input, int chunkBytes)
        {
            int rejected = 0;
            await Assert.ThrowsAsync<OperationCanceledException>(() => ChunkProducer.RunAsync(
                new MemoryStream(input), new XlsxRowCutter("<row"u8.ToArray(), "</sheetData"u8.ToArray()), chunkBytes, 0,
                chunk =>
                {
                    rejected = chunk.Buffer.GetHashCode();
                    return ValueTask.FromException(new OperationCanceledException("the consumer is gone"));
                },
                CancellationToken.None));
            return rejected;
        }

        [Fact]
        public async Task AChunkWhoseEmitFailsIsReturnedToThePool()
        {
            using PoolEvents events = new();

            int rejected = await RejectedEmitAsync(SheetXml(10_000), 4096);

            Assert.Contains(rejected, events.Returned);
        }

        [Fact]
        public async Task TheLastChunkIsReturnedToThePoolWhenItsEmitFails()
        {
            using PoolEvents events = new();

            int rejected = await RejectedEmitAsync(SheetXml(3), 4096);

            Assert.Contains(rejected, events.Returned);
        }

        private static readonly RowAggregation<long> Counting = new()
        {
            Seed = static () => 0,
            Accumulate = static (ref long count, Row _) => count++,
            Combine = static (left, right) => left + right,
        };

        private static async Task AssertEveryChunkReturnedAfterEarlyDisposeAsync(IExcelSheet sheet)
        {
            using PoolEvents events = new();
            SheetPipelineSource source = SheetPipelineSource.Create(sheet)!;
            await using (IAsyncEnumerator<long> results = SheetPipelineRunner.RunAsync(source, Counting, null, 0, 4, ChunkBytes, CancellationToken.None).GetAsyncEnumerator(TestContext.Current.CancellationToken))
            {
                Assert.True(await results.MoveNextAsync());
                await Task.Delay(50, TestContext.Current.CancellationToken);
            }

            Assert.NotEmpty(events.Outstanding);
            Assert.All(events.Outstanding, static pair => Assert.Equal(0, pair.Value));
        }

        [Fact]
        public async Task EveryXlsxChunkIsReturned_When_TheResultsAreAbandonedEarly()
        {
            using MemoryStream buffer = new();
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(buffer, leaveOpen: true))
            {
                WriteRows(workbook.AddSheet("S"));
                workbook.End();
            }
            buffer.Position = 0;
            using var reader = Excel.FromXlsx(buffer);

            await AssertEveryChunkReturnedAfterEarlyDisposeAsync(reader.FirstSheet);
        }

        [Fact]
        public async Task EveryXlsbChunkIsReturned_When_TheResultsAreAbandonedEarly()
        {
            using MemoryStream buffer = new();
            using (XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(buffer, leaveOpen: true))
            {
                WriteRows(workbook.AddSheet("S"));
                workbook.End();
            }
            buffer.Position = 0;
            using var reader = Excel.FromXlsb(buffer);

            await AssertEveryChunkReturnedAfterEarlyDisposeAsync(reader.FirstSheet);
        }

        private static void WriteRows<TRow>(Core.Writer.ISheetWriter<TRow> sheet)
            where TRow : Core.Writer.IRowWriter
        {
            using (sheet)
            {
                for (int r = 0; r < 200_000; r++)
                {
                    using TRow row = sheet.StartRow();
                    row.Write(string.Create(CultureInfo.InvariantCulture, $"item-{r:D7}"));
                    row.Write(r);
                    row.Write(r * 0.5);
                }
                sheet.End();
            }
        }
    }
}
