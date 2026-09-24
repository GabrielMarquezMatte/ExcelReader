using System.Runtime.InteropServices;
using BenchmarkDotNet.Attributes;
using ExcelReader.Native;
using ExcelReader.Native.Arrow;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Typed;

namespace ExcelReader.Benchmarks
{
    // NOTE: this measures MANAGED allocation only. The native blocks the columns are marshalled
    [MemoryDiagnoser]
    public class NativeTypedParseBenchmark
    {
        private const int HeaderRow = 1;
        private const long ExpectedRows = 65535;

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
            _xlsb = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Data", "65K_Records_Data.xlsb"));
        }

        [Benchmark(Baseline = true)]
        public long ParseTyped()
        {
            using NativeHandle handle = Open();
            int status = TypedApi.ParseTyped(handle, Schema, HeaderRow, out NativeTable table);
            try
            {
                VerifyParsed(status, table);
                return table.RowCount;
            }
            finally
            {
                TypedApi.FreeTable(ref table);
            }
        }

        [Benchmark]
        public long ParseArrow()
        {
            using NativeHandle handle = Open();
            int status = ArrowApi.ParseArrow(handle, Schema, HeaderRow, out ArrowArray array, out ArrowSchema schema);
            if (status != NativeStatus.Ok)
            {
                throw new InvalidOperationException($"xl_parse_arrow failed with status {status}.");
            }
            try
            {
                if (array.Length != ExpectedRows || array.NChildren != Schema.Length)
                {
                    throw new InvalidOperationException(
                        $"expected {ExpectedRows} rows x {Schema.Length} columns, got {array.Length} x {array.NChildren}.");
                }
                return array.Length;
            }
            finally
            {
                Release(array, schema);
            }
        }

        private static void Release(ArrowArray array, ArrowSchema schema)
        {
            IntPtr arrayPtr = Marshal.AllocHGlobal(Marshal.SizeOf<ArrowArray>());
            IntPtr schemaPtr = Marshal.AllocHGlobal(Marshal.SizeOf<ArrowSchema>());
            try
            {
                Marshal.StructureToPtr(array, arrayPtr, false);
                Marshal.StructureToPtr(schema, schemaPtr, false);
                ArrowApi.ReleaseArrowArray(arrayPtr);
                ArrowApi.ReleaseArrowSchema(schemaPtr);
            }
            finally
            {
                Marshal.FreeHGlobal(arrayPtr);
                Marshal.FreeHGlobal(schemaPtr);
            }
        }

        private static NativeColumnSpec Spec(string name, int type)
        {
            return new NativeColumnSpec { Names = [name], Type = type, Nullable = false };
        }

        private NativeHandle Open()
        {
            int status = ReadApi.OpenMemory(_xlsb, NativeFormat.Xlsb, out NativeHandle? handle);
            if (status != NativeStatus.Ok || handle is null)
            {
                throw new InvalidOperationException($"opening the xlsb fixture failed with status {status}.");
            }
            return handle;
        }

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
