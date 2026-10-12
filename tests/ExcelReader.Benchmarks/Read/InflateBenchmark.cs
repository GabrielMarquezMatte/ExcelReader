using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using ExcelReader.Core.Reader.Zip.Inflate;

namespace ExcelReader.Benchmarks
{
    [MemoryDiagnoser]
    [GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
    [CategoriesColumn]
    public class InflateBenchmark
    {
        private readonly byte[] _sink = new byte[64 * 1024];
        private byte[] _xlsxSheet = [];
        private byte[] _xlsbSheet = [];
        private byte[] _sharedStrings = [];

        [GlobalSetup]
        public async Task SetupAsync()
        {
            string dir = Path.Combine(AppContext.BaseDirectory, "Data");
            _xlsxSheet = RawDeflateEntry(await File.ReadAllBytesAsync(Path.Combine(dir, "65K_Records_Data.xlsx")), "xl/worksheets/sheet1.xml");
            _xlsbSheet = RawDeflateEntry(await File.ReadAllBytesAsync(Path.Combine(dir, "65K_Records_Data.xlsb")), "xl/worksheets/sheet1.bin");
            _sharedStrings = RawDeflateEntry(await StringHeavyWorkbookGenerator.BuildXlsxAsync(65_536), "xl/sharedStrings.xml");
            RequireSameOutput(_xlsxSheet);
            RequireSameOutput(_xlsbSheet);
            RequireSameOutput(_sharedStrings);
        }

        [Benchmark(Baseline = true)]
        [BenchmarkCategory("XlsxSheet")]
        public long XlsxSheet_DeflateStream()
        {
            return Drain(new DeflateStream(new MemoryStream(_xlsxSheet, writable: false), CompressionMode.Decompress));
        }

        [Benchmark]
        [BenchmarkCategory("XlsxSheet")]
        public long XlsxSheet_InflateStream()
        {
            return Drain(new InflateStream(new MemoryStream(_xlsxSheet, writable: false)));
        }

        [Benchmark(Baseline = true)]
        [BenchmarkCategory("XlsbSheet")]
        public long XlsbSheet_DeflateStream()
        {
            return Drain(new DeflateStream(new MemoryStream(_xlsbSheet, writable: false), CompressionMode.Decompress));
        }

        [Benchmark]
        [BenchmarkCategory("XlsbSheet")]
        public long XlsbSheet_InflateStream()
        {
            return Drain(new InflateStream(new MemoryStream(_xlsbSheet, writable: false)));
        }

        [Benchmark(Baseline = true)]
        [BenchmarkCategory("SharedStrings")]
        public long SharedStrings_DeflateStream()
        {
            return Drain(new DeflateStream(new MemoryStream(_sharedStrings, writable: false), CompressionMode.Decompress));
        }

        [Benchmark]
        [BenchmarkCategory("SharedStrings")]
        public long SharedStrings_InflateStream()
        {
            return Drain(new InflateStream(new MemoryStream(_sharedStrings, writable: false)));
        }

        private long Drain(Stream stream)
        {
            using (stream)
            {
                long total = 0;
                int read;
                while ((read = stream.Read(_sink)) > 0)
                {
                    total += read;
                }
                return total;
            }
        }

        private static void RequireSameOutput(byte[] deflated)
        {
            using var expected = new MemoryStream();
            using (var deflate = new DeflateStream(new MemoryStream(deflated, writable: false), CompressionMode.Decompress))
            {
                deflate.CopyTo(expected);
            }
            using var actual = new MemoryStream();
            using (var inflate = new InflateStream(new MemoryStream(deflated, writable: false)))
            {
                inflate.CopyTo(actual);
            }
            if (expected.Length == 0 || !expected.GetBuffer().AsSpan(0, (int)expected.Length).SequenceEqual(actual.GetBuffer().AsSpan(0, (int)actual.Length)))
            {
                throw new InvalidOperationException("InflateStream and DeflateStream disagree on a benchmark entry, or the entry is empty.");
            }
        }

        private static byte[] RawDeflateEntry(byte[] zip, string entryName)
        {
            int end = zip.AsSpan().LastIndexOf("PK\x05\x06"u8);
            int count = BinaryPrimitives.ReadUInt16LittleEndian(zip.AsSpan(end + 10));
            int position = (int)BinaryPrimitives.ReadUInt32LittleEndian(zip.AsSpan(end + 16));
            for (int i = 0; i < count; i++)
            {
                ReadOnlySpan<byte> header = zip.AsSpan(position);
                int nameLength = BinaryPrimitives.ReadUInt16LittleEndian(header[28..]);
                if (string.Equals(Encoding.UTF8.GetString(header.Slice(46, nameLength)), entryName, StringComparison.Ordinal))
                {
                    return ReadEntryData(zip, header);
                }
                position += 46 + nameLength + BinaryPrimitives.ReadUInt16LittleEndian(header[30..]) + BinaryPrimitives.ReadUInt16LittleEndian(header[32..]);
            }
            throw new InvalidOperationException($"Entry '{entryName}' not found.");
        }

        private static byte[] ReadEntryData(byte[] zip, ReadOnlySpan<byte> centralHeader)
        {
            if (BinaryPrimitives.ReadUInt16LittleEndian(centralHeader[10..]) != 8)
            {
                throw new InvalidOperationException("The benchmark entry is not deflated.");
            }
            int compressedSize = (int)BinaryPrimitives.ReadUInt32LittleEndian(centralHeader[20..]);
            int local = (int)BinaryPrimitives.ReadUInt32LittleEndian(centralHeader[42..]);
            ReadOnlySpan<byte> localHeader = zip.AsSpan(local);
            int data = local + 30 + BinaryPrimitives.ReadUInt16LittleEndian(localHeader[26..]) + BinaryPrimitives.ReadUInt16LittleEndian(localHeader[28..]);
            return zip.AsSpan(data, compressedSize).ToArray();
        }
    }
}
