using BenchmarkDotNet.Attributes;
using ExcelReader.Native;

namespace ExcelReader.Benchmarks
{
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
