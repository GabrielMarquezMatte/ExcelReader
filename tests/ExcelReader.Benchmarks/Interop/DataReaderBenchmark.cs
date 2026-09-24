using System.Data;
using System.Diagnostics.CodeAnalysis;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;

namespace ExcelReader.Benchmarks
{
    [MemoryDiagnoser]
    [SuppressMessage("Performance", "CA1859:Use concrete types when possible for improved performance",
        Justification = "The interface dispatch is the thing being measured: SqlBulkCopy, DataTable.Load and " +
            "Dapper all hold an IDataReader over an IExcelRowReader, so binding these to the concrete types " +
            "would devirtualize calls a real consumer pays for and report a cost nobody actually sees.")]
    public class DataReaderBenchmark
    {
        [Params(50_000)]
        public int Rows { get; set; }

        private byte[] _xlsx = [];
        private readonly byte[] _byteBuffer = new byte[256];

        [GlobalSetup]
        public async Task SetupAsync()
        {
            _xlsx = await WorkbookGenerator.BuildTypedAsync(Rows).ConfigureAwait(false);
        }

        private IExcelRowReader OpenReader()
        {
            return Excel.FromXlsx(new MemoryStream(_xlsx, writable: false), leaveOpen: false);
        }

        [Benchmark(Baseline = true)]
        public long Baseline_RawRows()
        {
            using IExcelRowReader reader = OpenReader();
            long acc = 0;
            foreach (Row row in reader)
            {
                for (int i = 0; i < row.ColumnCount; i++)
                {
                    acc += row[i].Value.Length;
                }
            }
            return acc;
        }

        [Benchmark]
        public long DataReader_GetValue()
        {
            using IExcelRowReader reader = OpenReader();
            using IDataReader data = new ExcelDataReader(reader);
            long acc = 0;
            while (data.Read())
            {
                for (int i = 0; i < data.FieldCount; i++)
                {
                    acc += data.GetValue(i) is null ? 0 : 1;
                }
            }
            return acc;
        }

        [Benchmark]
        public long DataReader_TypedGetters()
        {
            using IExcelRowReader reader = OpenReader();
            using IDataReader data = new ExcelDataReader(reader);
            long acc = 0;
            while (data.Read())
            {
                acc += data.GetString(0).Length;
                acc += data.GetInt32(1);
                acc += data.GetDateTime(2).Day;
                acc += (long)data.GetDouble(3);
            }
            return acc;
        }

        [Benchmark]
        public long DataReader_GetBytes()
        {
            using IExcelRowReader reader = OpenReader();
            using IDataReader data = new ExcelDataReader(reader);
            long acc = 0;
            while (data.Read())
            {
                acc += data.GetBytes(0, 0, _byteBuffer, 0, _byteBuffer.Length);
            }
            return acc;
        }

        [Benchmark]
        public int DataTable_Load()
        {
            using IExcelRowReader reader = OpenReader();
            using IDataReader data = new ExcelDataReader(reader);
            var table = new DataTable();
            table.Load(data);
            return table.Rows.Count;
        }
    }
}
