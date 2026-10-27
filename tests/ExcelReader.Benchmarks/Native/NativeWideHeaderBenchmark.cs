using System.Text;
using BenchmarkDotNet.Attributes;
using ExcelReader.Native;
using ExcelReader.Native.Reading;
using ExcelReader.Native.Typed;

namespace ExcelReader.Benchmarks
{
    // Binding every column of a wide header by name: the setup cost of a typed parse, before any data row.
    [MemoryDiagnoser]
    public class NativeWideHeaderBenchmark
    {
        [Params(512, 4096)]
        public int Columns { get; set; }

        private byte[] _csv = [];
        private NativeColumnSpec[] _specs = [];

        [GlobalSetup]
        public void Setup()
        {
            string[] names = [.. Enumerable.Range(0, Columns).Select(static c => $"column_{c:D5}")];
            _csv = Encoding.UTF8.GetBytes(string.Join(',', names) + "\n");
            _specs = [.. names.Select(static name => new NativeColumnSpec { Names = [name], Type = NativeColumnType.String })];
        }

        [Benchmark]
        public long BindByName()
        {
            if (ReadApi.OpenMemory(_csv, NativeFormat.Csv, out NativeHandle? handle) != NativeStatus.Ok || handle is null)
            {
                throw new InvalidOperationException("opening the header-only CSV failed.");
            }
            using (handle)
            {
                int status = TypedApi.ParseTyped(handle, 0, _specs, headerRow: 1, out NativeTable table);
                if (status != NativeStatus.Ok)
                {
                    throw new InvalidOperationException($"xl_parse_typed failed with status {status}.");
                }
                TypedApi.FreeTable(ref table);
                return table.ColumnCount;
            }
        }
    }
}
