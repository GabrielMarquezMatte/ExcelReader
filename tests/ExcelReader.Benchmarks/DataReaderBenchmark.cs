using System.Data;
using System.Diagnostics.CodeAnalysis;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Benchmarks
{
    // ExcelDataReader — the System.Data.IDataReader adapter — is the entry point for SqlBulkCopy,
    // DataTable.Load and Dapper, and had no benchmark at all. What the legs separate:
    //
    // - Baseline_RawRows is the same sheet read straight through IExcelRowReader, so the adapter's
    //   own cost is the gap between it and everything below, not an absolute number.
    // - GetValue is how SqlBulkCopy actually drives a reader: one boxed object per cell, which is
    //   the allocation floor of any bulk load through this adapter.
    // - TypedGetters is the same rows read by a consumer that already knows its schema, and is the
    //   honest comparison for "what does the adapter cost", since it skips the boxing GetValue
    //   cannot avoid.
    // - GetBytes is the binary-column path on a bulk load: it reads Cell.Value's UTF-8 straight out
    //   with no string in between, so it should not allocate per cell.
    // - DataTable_Load is the other headline consumer, and carries DataTable's own per-row cost —
    //   it is here to size that against the adapter, not to measure the adapter alone.
    //
    // The corpus is WorkbookGenerator.BuildTypedAsync: a header row plus Name/Id/Date/Value, so the
    // four columns exercise the string, integer, date and floating-point getters respectively.
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
            // leaveOpen: false so disposing the reader disposes the stream with it.
            return Excel.FromXlsx(new MemoryStream(_xlsx, writable: false), leaveOpen: false);
        }

        // What the sheet costs with no adapter in the way.
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

        // SqlBulkCopy's access pattern: an untyped, boxed read per cell.
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

        // A consumer that already knows the schema, so nothing boxes.
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

        // The binary-column path: straight out of the cell's UTF-8, no string round-trip.
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

        // Carries DataTable's own per-row cost on top of the adapter's.
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
