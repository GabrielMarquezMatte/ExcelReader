using BenchmarkDotNet.Attributes;
using ExcelReader.Native;

namespace ExcelReader.Benchmarks
{
    // The allocation counterpart to python/benchmarks/bench_read.py's parse_typed timing. That
    // harness measures the whole native round trip in wall clock, where the columnar build phase is
    // a small enough slice of a mostly-inflate-and-parse workload that a real change to it lands
    // inside the run-to-run noise. Allocation is deterministic, so it is the measurement that can
    // actually attribute a change to this path — the same reasoning STYLEGUIDE.md's "Tests and
    // Benchmarks" section gives for preferring allocation figures over timings.
    //
    // Drives NativeApi rather than the C ABI for the reason ExcelReader.Native's csproj states:
    // [UnmanagedCallersOnly] exports cannot be invoked from managed code.
    //
    // NOTE: this measures MANAGED allocation only. The native blocks the columns are marshalled
    // into (Marshal.AllocHGlobal) are invisible to MemoryDiagnoser by design — they are the output,
    // not the overhead. What moves here is the managed scratch spent producing them.
    [MemoryDiagnoser]
    public class NativeTypedParseBenchmark
    {
        // Row 1 of the fixture is its header; parse_typed consumes it, so the data row count is one
        // short of the file's 65536.
        private const int HeaderRow = 1;
        private const long ExpectedRows = 65535;

        // The 14 columns of Data/65K_Records_Data.xlsb in file order — deliberately the same schema
        // python/benchmarks/bench_read.py uses, so the timing harness and this one measure the same
        // work on the same file and their numbers can be read side by side.
        private static readonly NativeColumnSpec[] Schema =
        [
            Spec("Region", NativeColumnType.String),
            Spec("Country", NativeColumnType.String),
            Spec("Item Type", NativeColumnType.String),
            Spec("Sales Channel", NativeColumnType.String),
            Spec("Order Priority", NativeColumnType.String),
            Spec("Order Date", NativeColumnType.Date),
            Spec("Order ID", NativeColumnType.Int64),
            Spec("Ship Date", NativeColumnType.Date),
            Spec("Units Sold", NativeColumnType.Int64),
            Spec("Unit Price", NativeColumnType.Float64),
            Spec("Unit Cost", NativeColumnType.Float64),
            Spec("Total Revenue", NativeColumnType.Float64),
            Spec("Total Cost", NativeColumnType.Float64),
            Spec("Total Profit", NativeColumnType.Float64),
        ];

        private byte[] _xlsb = [];

        [GlobalSetup]
        public void Setup()
        {
            // ReadAllBytes throws if the fixture was never copied to the output directory, which is
            // the loud failure wanted here — a benchmark that silently ran on nothing would publish
            // an allocation figure indistinguishable from a large improvement.
            _xlsb = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Data", "65K_Records_Data.xlsb"));
        }

        [Benchmark]
        public long ParseTyped()
        {
            using NativeHandle handle = Open();
            int status = NativeApi.ParseTyped(handle, Schema, HeaderRow, out NativeTable table);
            try
            {
                VerifyParsed(status, table);
                return table.RowCount;
            }
            finally
            {
                NativeApi.FreeTable(ref table);
            }
        }

        private static NativeColumnSpec Spec(string name, int type)
        {
            return new NativeColumnSpec { Name = name, Type = type, Nullable = false };
        }

        private NativeHandle Open()
        {
            int status = NativeApi.OpenMemory(_xlsb, NativeFormat.Xlsb, out NativeHandle? handle);
            if (status != NativeStatus.Ok || handle is null)
            {
                throw new InvalidOperationException($"opening the xlsb fixture failed with status {status}.");
            }
            return handle;
        }

        // Asserted every iteration rather than once in setup: a parse that quietly yields nothing
        // still publishes an allocation number, and a number measured over no work reads exactly
        // like a win.
        private static void VerifyParsed(int status, NativeTable table)
        {
            if (status != NativeStatus.Ok)
            {
                throw new InvalidOperationException($"xl_parse_typed failed with status {status}.");
            }
            if (table.RowCount != ExpectedRows || table.ColumnCount != Schema.Length)
            {
                throw new InvalidOperationException(
                    $"expected {ExpectedRows} rows x {Schema.Length} columns, got {table.RowCount} x {table.ColumnCount}.");
            }
        }
    }
}
