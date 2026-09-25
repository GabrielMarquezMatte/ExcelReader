using System.Buffers.Text;
using System.IO.Compression;
using System.Text.RegularExpressions;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;
using static ExcelReader.Benchmarks.BenchmarkAccumulators;

namespace ExcelReader.Benchmarks
{
    // Sheet stage = full read minus the shared-string prologue; the Stored_* pair repeats both with
    // the sheet re-zipped uncompressed, so their difference is sheet parse with no inflate floor.
    [MemoryDiagnoser]
    public partial class XlsxSharedStringHotPathBenchmark
    {
        [Params(65_536)]
        public int Rows { get; set; }

        private static readonly ExcelReaderOptions _prefetchOptions = new() { PrefetchDecompression = true };

        private byte[] _xlsx = [];
        private byte[] _stored = [];
        private byte[][] _indices = [];

        [GeneratedRegex("t=\"s\"><v>(?<index>[0-9]+)</v>", RegexOptions.NonBacktracking)]
        private static partial Regex SharedIndexPattern();

        [GlobalSetup]
        public async Task SetupAsync()
        {
            _xlsx = await StringHeavyWorkbookGenerator.BuildXlsxAsync(Rows);
            _stored = Restore(_xlsx);
            using var zip = new ZipArchive(new MemoryStream(_xlsx, writable: false), ZipArchiveMode.Read);
            using var sheet = new StreamReader(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open());
            string xml = await sheet.ReadToEndAsync();
            _indices = [.. SharedIndexPattern().Matches(xml).Select(m => System.Text.Encoding.ASCII.GetBytes(m.Groups["index"].Value))];
        }

        private static byte[] Restore(byte[] source)
        {
            using var src = new ZipArchive(new MemoryStream(source, writable: false), ZipArchiveMode.Read);
            using var ms = new MemoryStream();
            using (var dst = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (ZipArchiveEntry entry in src.Entries)
                {
                    using Stream input = entry.Open();
                    using Stream output = dst.CreateEntry(entry.FullName, CompressionLevel.NoCompression).Open();
                    input.CopyTo(output);
                }
            }
            return ms.ToArray();
        }

        private static long ReadAll(byte[] workbook, ExcelReaderOptions? options)
        {
            using var ms = new MemoryStream(workbook, writable: false);
            using XlsxReader reader = Excel.FromXlsx(ms, options: options);
            long acc = 0;
            foreach (Row row in reader) { acc += AccumulateRow(row); }
            return acc;
        }

        private static bool Prologue(byte[] workbook)
        {
            using var ms = new MemoryStream(workbook, writable: false);
            using XlsxReader reader = Excel.FromXlsx(ms);
            using XlsxReader.Enumerator rows = reader.GetEnumerator();
            return rows.MoveNext();
        }

        [Benchmark(Baseline = true)]
        public long EndToEnd() => ReadAll(_xlsx, null);

        [Benchmark]
        public long EndToEnd_Prefetch() => ReadAll(_xlsx, _prefetchOptions);

        [Benchmark]
        public bool SharedStringPrologue() => Prologue(_xlsx);

        [Benchmark]
        public long Stored_EndToEnd() => ReadAll(_stored, null);

        [Benchmark]
        public bool Stored_SharedStringPrologue() => Prologue(_stored);

        [Benchmark]
        public long IndexParse_Utf8Parser()
        {
            long acc = 0;
            foreach (byte[] index in _indices)
            {
                if (Utf8Parser.TryParse(index, out int value, out _))
                {
                    acc += value;
                }
            }
            return acc;
        }

        [Benchmark]
        public long IndexParse_DigitLoop()
        {
            long acc = 0;
            foreach (byte[] index in _indices)
            {
                int value = 0;
                foreach (byte b in index)
                {
                    value = (value * 10) + (b - '0');
                }
                acc += value;
            }
            return acc;
        }
    }
}
