using System.Buffers;
using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser.SheetPipeline;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests.Parser.SheetPipeline
{
    public class ChunkProducerTests
    {
        private static byte[] Sheet(int rows, int padding = 0)
        {
            var sb = new StringBuilder("<worksheet><sheetData>");
            for (int r = 1; r <= rows; r++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"<row r=\"{r}\"><c><v>{r}</v></c>{new string(' ', padding)}</row>");
            }
            return Encoding.UTF8.GetBytes(sb.Append("</sheetData></worksheet>").ToString());
        }

        private static async Task<List<byte[]>> ProduceAsync(byte[] input, int chunkBytes, int maxCellBytes = 0, RowCutter? cutter = null)
        {
            var chunks = new List<byte[]>();
            await ChunkProducer.RunAsync(
                new MemoryStream(input), cutter ?? new XlsxRowCutter("<row"u8.ToArray()), chunkBytes, maxCellBytes,
                chunk =>
                {
                    chunks.Add(chunk.Buffer.AsSpan(0, chunk.Length).ToArray());
                    ArrayPool<byte>.Shared.Return(chunk.Buffer);
                    return ValueTask.CompletedTask;
                },
                TestContext.Current.CancellationToken);
            return chunks;
        }

        [Theory]
        [InlineData(64)]
        [InlineData(100)]
        [InlineData(4096)]
        public async Task ChunksConcatenateToTheInputAndStartOnRows(int chunkBytes)
        {
            byte[] input = Sheet(200);

            List<byte[]> chunks = await ProduceAsync(input, chunkBytes);

            Assert.Equal(input, chunks.SelectMany(c => c).ToArray());
            Assert.True(chunks.Count > 1);
            Assert.All(chunks.Skip(1), chunk => Assert.StartsWith("<row ", Encoding.UTF8.GetString(chunk), StringComparison.Ordinal));
        }

        [Fact]
        public async Task AnEmptyStreamProducesNothing()
        {
            Assert.Empty(await ProduceAsync([], 64));
        }

        [Fact]
        public async Task ARowLongerThanTheChunkGrowsTheBuffer()
        {
            byte[] input = Sheet(5, padding: 500);

            List<byte[]> chunks = await ProduceAsync(input, 64);

            Assert.Equal(input, chunks.SelectMany(c => c).ToArray());
            Assert.All(chunks.Skip(1), chunk => Assert.StartsWith("<row ", Encoding.UTF8.GetString(chunk), StringComparison.Ordinal));
        }

        [Fact]
        public async Task GrowthStopsAtMaxCellBytes()
        {
            byte[] input = Sheet(5, padding: 5000);

            await Assert.ThrowsAsync<ExcelLimitExceededException>(() => ProduceAsync(input, 64, maxCellBytes: 1024));
        }

        [Fact]
        public async Task ARowLikeNameIsNotACut()
        {
            byte[] input = Encoding.UTF8.GetBytes("<sheetData><row r=\"1\"><c><v>1</v></c></row><rowBreaks count=\"1\"/><rowBreaks count=\"2\"/><rowBreaks count=\"3\"/></sheetData>");

            List<byte[]> chunks = await ProduceAsync(input, 48);

            Assert.Equal(input, chunks.SelectMany(c => c).ToArray());
            Assert.All(chunks.Skip(1), chunk => Assert.StartsWith("<row ", Encoding.UTF8.GetString(chunk), StringComparison.Ordinal));
        }

        private static byte[] Record(int id, int payload)
        {
            byte[] record = new byte[2 + payload];
            record[0] = (byte)id;
            record[1] = (byte)payload;
            return record;
        }

        [Fact]
        public async Task XlsbChunksStartOnRowHeadersAndStopAtEndSheetData()
        {
            const int Cell = 7;
            const int EndSheetData = 146;
            var input = new List<byte>();
            for (int r = 0; r < 40; r++)
            {
                input.AddRange(Record(0, 8));
                input.AddRange(Record(Cell, 12));
                input.AddRange(Record(Cell, 12));
            }
            byte[] end = [(byte)((EndSheetData & 0x7F) | 0x80), (byte)(EndSheetData >> 7), 0];
            input.AddRange(end);
            int dataLength = input.Count;
            input.AddRange(Record(0, 8));

            List<byte[]> chunks = await ProduceAsync([.. input], 96, cutter: new XlsbRowCutter());

            Assert.Equal(input.Take(dataLength), chunks.SelectMany(c => c));
            Assert.True(chunks.Count > 1);
            Assert.All(chunks, chunk => Assert.Equal(0, chunk[0]));
        }
    }
}
