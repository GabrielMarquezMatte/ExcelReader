using BenchmarkDotNet.Attributes;
using ExcelReader.Native;

namespace ExcelReader.Benchmarks
{
    // What this DOES measure: cumulative managed allocation (BenchmarkDotNet's Allocated column)
    // and wall-clock time, both summed/averaged over the whole call.
    //
    // What this does NOT measure, and cannot: the memory ceiling this feature actually exists for -
    // the peak bytes live at any one instant. Allocated is a running total across every batch, not
    // a high-water mark, and it is managed-heap-only: the Marshal.AllocHGlobal blocks that hold the
    // actual column data (the bulk of what the chunked ABI bounds) are invisible to it by design,
    // same as NativeTypedParseBenchmark's own note on this. A benchmark built on MemoryDiagnoser
    // cannot show the ceiling holding even if it does.
    //
    // Measured outcome (shortened run: --warmupCount 3 --iterationCount 5, so Mean/Error here carry
    // real run-to-run margin - Allocated does not need many iterations to be trustworthy):
    // WholeSheet 1.33 MB, Batched_10k 1.63 MB (1.23x), Batched_1k 1.56 MB (1.17x). Batching costs
    // MORE total allocation and, at this iteration count, no less time - a real and correct result,
    // not a bug: TypedParseSession.NextBatch allocates a fresh ColumnBuilder[]/ChunkedBuffer<T>
    // chain from empty on every batch (that restart is exactly what keeps any one batch's peak
    // memory low), so more batches means more restarts, which raises the running total even as the
    // peak footprint falls. Cumulative allocation and peak memory are different quantities that
    // move in opposite directions here.
    //
    // The ceiling itself - the peak, bytes-live-at-once quantity this feature actually claims - is
    // proven deterministically instead, by
    // TypedParseSessionTests.NextBatch_Should_Bound_The_Largest_Live_Table_To_Roughly_One_Batch in
    // tests/ExcelReader.Tests/TypedParseSessionTests.cs, which sums each live NativeTable's owned
    // bytes directly from its column descriptors and asserts the largest batch stays close to
    // maxRows/totalRows of the unbounded table's size. Read that test for the ceiling; read this
    // benchmark only for the allocation/time trade-off batching costs.
    [MemoryDiagnoser]
    public class ChunkedParseBenchmark
    {
        [Params(50_000)]
        public int Rows { get; set; }

        private string _path = "";

        private static NativeColumnSpec[] Specs()
        {
            return
            [
                new() { Names = ["Name"], Type = NativeColumnType.String, Nullable = true },
                new() { Names = ["Id"], Type = NativeColumnType.Int64, Nullable = true },
                new() { Names = ["Value"], Type = NativeColumnType.Float64, Nullable = true },
            ];
        }

        [GlobalSetup]
        public async Task SetupAsync()
        {
            _path = Path.Combine(Path.GetTempPath(), $"chunked-parse-{Rows}.xlsx");
            await File.WriteAllBytesAsync(_path, await WorkbookGenerator.BuildTypedAsync(Rows).ConfigureAwait(false))
                .ConfigureAwait(false);
        }

        [GlobalCleanup]
        public void Cleanup()
        {
            File.Delete(_path);
        }

        private NativeHandle OpenWorkbook()
        {
            int status = NativeApi.OpenFile(System.Text.Encoding.UTF8.GetBytes(_path), NativeFormat.Auto,
                out NativeHandle? handle);
            return status == NativeStatus.Ok && handle is not null
                ? handle
                : throw new InvalidOperationException($"open failed with status {status}");
        }

        [Benchmark(Baseline = true)]
        public long WholeSheet()
        {
            using NativeHandle handle = OpenWorkbook();
            int status = NativeApi.ParseTyped(handle, Specs(), headerRow: 1, out NativeTable table);
            if (status != NativeStatus.Ok)
            {
                throw new InvalidOperationException($"parse failed with status {status}");
            }
            try
            {
                return table.RowCount;
            }
            finally
            {
                NativeApi.FreeTable(ref table);
            }
        }

        private long Drain(long maxRows)
        {
            using NativeHandle handle = OpenWorkbook();
            int status = NativeApi.OpenTypedReader(handle, Specs(), headerRow: 1, maxRows, out nint reader);
            if (status != NativeStatus.Ok)
            {
                throw new InvalidOperationException($"open reader failed with status {status}");
            }
            try
            {
                long rows = 0;
                while (true)
                {
                    status = NativeApi.NextTypedBatch(reader, out NativeTable table);
                    if (status == NativeStatus.Eof)
                    {
                        return rows;
                    }
                    if (status != NativeStatus.Ok)
                    {
                        throw new InvalidOperationException($"next batch failed with status {status}");
                    }
                    try
                    {
                        rows += table.RowCount;
                    }
                    finally
                    {
                        // Freed as each batch is consumed - this is what keeps the ceiling flat.
                        NativeApi.FreeTable(ref table);
                    }
                }
            }
            finally
            {
                NativeApi.CloseTypedReader(reader);
            }
        }

        [Benchmark]
        public long Batched_10k()
        {
            return Drain(10_000);
        }

        [Benchmark]
        public long Batched_1k()
        {
            return Drain(1_000);
        }
    }
}
