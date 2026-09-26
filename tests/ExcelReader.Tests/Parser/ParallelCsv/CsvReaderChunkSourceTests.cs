using System.Text;
using ExcelReader.Core.Parser.ParallelCsv;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Tests.Parser.ParallelCsv
{
    public class CsvReaderChunkSourceTests
    {
        private static byte[] Numbered(int rows)
        {
            StringBuilder text = new("n\n");
            for (int i = 0; i < rows; i++)
            {
                text.Append(i).Append('\n');
            }
            return Encoding.UTF8.GetBytes(text.ToString());
        }

        [Fact]
        public void TryGetChunkSource_Should_Expose_A_Memory_Reader_Buffer()
        {
            byte[] csv = Numbered(10);
            using CsvReader reader = Excel.FromCsv(csv.AsMemory());

            Assert.True(reader.TryGetChunkSource(out CsvChunkSource source));
            Assert.Equal(csv.Length, source.Length);
            Assert.True(source.IsMemory);
        }

        [Fact]
        public void TryGetChunkSource_Should_Expose_A_File_Reader_Handle()
        {
            string path = Path.Combine(Path.GetTempPath(), $"excelreader-chunk-{Guid.NewGuid():N}.csv");
            File.WriteAllBytes(path, Numbered(10));
            try
            {
                using CsvReader reader = Excel.FromCsvFile(path);

                Assert.True(reader.TryGetChunkSource(out CsvChunkSource source));
                Assert.Equal(new FileInfo(path).Length, source.Length);
                Assert.False(source.IsMemory);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void TryGetChunkSource_Should_Refuse_A_Stream_That_Is_Not_A_File()
        {
            using CsvReader reader = Excel.FromCsv(new MemoryStream(Numbered(10)), leaveOpen: false);

            Assert.False(reader.TryGetChunkSource(out _));
        }

        [Fact]
        public async Task RunPartitionedAsync_Should_Deliver_Every_Record_After_The_Header_Once()
        {
            byte[] csv = Numbered(50_000);
            using CsvReader reader = Excel.FromCsv(csv.AsMemory());
            Assert.True(reader.TryGetChunkSource(out CsvChunkSource source));
            CsvAggregation<long> count = new()
            {
                Seed = () => 0L,
                Accumulate = (ref long seen, Row _) => seen++,
                Combine = (left, right) => left + right,
            };
            CsvParallelOptions options = new() { DegreeOfParallelism = 4, HeaderRow = 1, Reader = reader.Options };

            long rows = await ParallelCsvProcessor.RunPartitionedAsync(source, count, options, chunkSizeOverride: 4096, CancellationToken.None);

            Assert.Equal(50_000, rows);
        }
    }
}
