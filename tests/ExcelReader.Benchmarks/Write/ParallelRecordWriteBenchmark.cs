using System.Diagnostics.CodeAnalysis;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Benchmarks
{
    [MemoryDiagnoser]
    [GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
    [CategoriesColumn]
    [SuppressMessage("Design", "CA1001", Justification = "BenchmarkDotNet owns the instance; GlobalCleanup releases the stream.")]
    public class ParallelRecordWriteBenchmark
    {
        private const int Rows = 1_000_000;

        private static readonly ExcelRecordLayout<Record> Layout = ExcelRecordLayout.FromAttributes<Record>();

        private readonly MemoryStream _output = new(32 * 1024 * 1024);
        private List<Record> _records = [];

        [GlobalSetup]
        public void Setup()
        {
            _records = WorkbookGenerator.Records(Rows);
        }

        [GlobalCleanup]
        public void Cleanup()
        {
            _output.Dispose();
        }

        private async Task<long> WriteXlsxAsync(int degree, bool prefetch = false)
        {
            _output.SetLength(0);
            await using (var writer = XlsxWorkbookWriter.Create(_output, leaveOpen: true, options: new XlsxWriterOptions { PrefetchWrite = prefetch }))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await WriteAsync(sheet, degree);
                }
            }
            return _output.Length;
        }

        private async Task<long> WriteXlsbAsync(int degree, bool prefetch = false)
        {
            _output.SetLength(0);
            await using (var writer = XlsbWorkbookWriter.Create(_output, leaveOpen: true, options: new XlsbWriterOptions { PrefetchWrite = prefetch }))
            {
                await using (var sheet = writer.AddSheet("S1"))
                {
                    await WriteAsync(sheet, degree);
                }
            }
            return _output.Length;
        }

        private ValueTask WriteAsync(XlsxSheetWriter sheet, int degree)
        {
            if (degree == 0)
            {
                return sheet.WriteRecordsAsync(_records, Layout);
            }
            return sheet.WriteRecordsParallelAsync(_records, Layout, new ParallelWriteOptions { DegreeOfParallelism = degree });
        }

        private ValueTask WriteAsync(XlsbSheetWriter sheet, int degree)
        {
            if (degree == 0)
            {
                return sheet.WriteRecordsAsync(_records, Layout);
            }
            return sheet.WriteRecordsParallelAsync(_records, Layout, new ParallelWriteOptions { DegreeOfParallelism = degree });
        }

        [Benchmark(Baseline = true)]
        [BenchmarkCategory("Xlsx")]
        public Task<long> Xlsx_Serial()
        {
            return WriteXlsxAsync(0);
        }

        [Benchmark]
        [BenchmarkCategory("Xlsx")]
        public Task<long> Xlsx_Serial_Prefetch()
        {
            return WriteXlsxAsync(0, prefetch: true);
        }

        [Benchmark]
        [BenchmarkCategory("Xlsx")]
        public Task<long> Xlsx_Parallel_1()
        {
            return WriteXlsxAsync(1);
        }

        [Benchmark]
        [BenchmarkCategory("Xlsx")]
        public Task<long> Xlsx_Parallel_2()
        {
            return WriteXlsxAsync(2);
        }

        [Benchmark]
        [BenchmarkCategory("Xlsx")]
        public Task<long> Xlsx_Parallel_4()
        {
            return WriteXlsxAsync(4);
        }

        [Benchmark]
        [BenchmarkCategory("Xlsx")]
        public Task<long> Xlsx_Parallel_8()
        {
            return WriteXlsxAsync(8);
        }

        [Benchmark]
        [BenchmarkCategory("Xlsx")]
        public Task<long> Xlsx_Parallel_16()
        {
            return WriteXlsxAsync(16);
        }

        [Benchmark(Baseline = true)]
        [BenchmarkCategory("Xlsb")]
        public Task<long> Xlsb_Serial()
        {
            return WriteXlsbAsync(0);
        }

        [Benchmark]
        [BenchmarkCategory("Xlsb")]
        public Task<long> Xlsb_Serial_Prefetch()
        {
            return WriteXlsbAsync(0, prefetch: true);
        }

        [Benchmark]
        [BenchmarkCategory("Xlsb")]
        public Task<long> Xlsb_Parallel_1()
        {
            return WriteXlsbAsync(1);
        }

        [Benchmark]
        [BenchmarkCategory("Xlsb")]
        public Task<long> Xlsb_Parallel_2()
        {
            return WriteXlsbAsync(2);
        }

        [Benchmark]
        [BenchmarkCategory("Xlsb")]
        public Task<long> Xlsb_Parallel_4()
        {
            return WriteXlsbAsync(4);
        }

        [Benchmark]
        [BenchmarkCategory("Xlsb")]
        public Task<long> Xlsb_Parallel_8()
        {
            return WriteXlsbAsync(8);
        }

        [Benchmark]
        [BenchmarkCategory("Xlsb")]
        public Task<long> Xlsb_Parallel_16()
        {
            return WriteXlsbAsync(16);
        }
    }
}
