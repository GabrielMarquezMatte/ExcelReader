using System.Globalization;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Tests
{
    public class CsvChunkWorkerTests
    {
        private sealed class Row
        {
            public string? Name { get; set; }
            public int Age { get; set; }
        }

        private static async Task<CsvChunkResult<Row>> ParseAsync(byte[] csv, long start, long end, long? confirmedStart = null)
        {
            using var headerReader = Excel.FromCsv(csv);
            CsvBoundColumnMap<Row> map = CsvHeaderBinder.Bind<Row>(
                headerReader, new ExcelParserConfig(), TypeMapper<Row>.GetCsvInfo(), out _);

            var source = new CsvChunkSource(csv.AsMemory());
            return await CsvChunkWorker.ParseAsync(
                source,
                new CsvChunk(0, start, end),
                confirmedStart,
                map,
                TypeMapper<Row>.GetCsvInfo(),
                CsvReaderOptions.Default,
                new ExcelParserConfig(),
                CancellationToken.None);
        }

        [Fact]
        public async Task ParsesOnlyTheRecordsStartingInsideItsChunk()
        {
            //                  0        9        18       27
            byte[] csv = "Name,Age\nAda,0001\nBob,0002\nCid,0003\n"u8.ToArray();

            // Chunk covering just the "Bob" record.
            CsvChunkResult<Row> result = await ParseAsync(csv, start: 18, end: 27, confirmedStart: 18);

            Assert.Single(result.Models);
            Assert.Equal("Bob", result.Models[0].Name);
        }

        [Fact]
        public async Task FinishesTheRecordThatStraddlesTheChunkEnd()
        {
            byte[] csv = "Name,Age\nAda,0001\nBob,0002\nCid,0003\n"u8.ToArray();

            // End lands in the middle of the "Bob" record, which started inside the chunk.
            CsvChunkResult<Row> result = await ParseAsync(csv, start: 9, end: 22, confirmedStart: 9);

            Assert.Equal(2, result.Models.Count);
            Assert.Equal("Ada", result.Models[0].Name);
            Assert.Equal("Bob", result.Models[1].Name);
            Assert.Equal(2, result.Models[1].Age);
        }

        [Fact]
        public async Task ReportsTheFirstRecordStartAtOrAfterItsEndAsResolvedNextStart()
        {
            byte[] csv = "Name,Age\nAda,0001\nBob,0002\nCid,0003\n"u8.ToArray();

            CsvChunkResult<Row> result = await ParseAsync(csv, start: 9, end: 22, confirmedStart: 9);

            Assert.Equal(27, result.ResolvedNextStart);
        }

        [Fact]
        public async Task ResolvedNextStartIsMaxValueWhenTheChunkRunsToEof()
        {
            byte[] csv = "Name,Age\nAda,0001\n"u8.ToArray();

            CsvChunkResult<Row> result = await ParseAsync(csv, start: 9, end: 18, confirmedStart: 9);

            Assert.Equal(long.MaxValue, result.ResolvedNextStart);
        }

        [Fact]
        public async Task GuessesItsOwnStartWhenNoConfirmedStartIsGiven()
        {
            byte[] csv = "Name,Age\nAda,0001\nBob,0002\nCid,0003\n"u8.ToArray();

            // Start lands mid-record (inside "Ada,0001"); the resolver must skip to the next record.
            CsvChunkResult<Row> result = await ParseAsync(csv, start: 12, end: 27);

            Assert.Equal(18, result.ActualStart);
            Assert.Single(result.Models);
            Assert.Equal("Bob", result.Models[0].Name);
        }

        [Fact]
        public async Task GuessesWrongWhenANewlineHidesInsideAQuotedField()
        {
            // The \n inside the quoted "a\nb" is not a record boundary, but a chunk starting
            // mid-field cannot know that. This test pins the *wrong* guess so the merge in Task 7
            // has something concrete to correct.
            byte[] csv = "Name,Age\n\"a\nb\",0001\nCid,0003\n"u8.ToArray();

            CsvChunkResult<Row> result = await ParseAsync(csv, start: 10, end: 20);

            // The Outside hypothesis takes the \n at index 11 as a terminator, so it starts at 12.
            Assert.Equal(12, result.ActualStart);
        }

        [Fact]
        public async Task YieldsNothingWhenTheChunkContainsNoRecordStart()
        {
            // One enormous quoted field spanning the whole chunk: no record starts inside it.
            byte[] csv = "Name,Age\n\"aaaaaaaaaaaaaaaaaaaaaaaaaaaa\",1\n"u8.ToArray();

            CsvChunkResult<Row> result = await ParseAsync(csv, start: 15, end: 25);

            Assert.Empty(result.Models);
        }

        [Fact]
        public async Task ParsesTheSameRecordFromAFileBackedSourceAsFromMemory()
        {
            //                  0        9        18       27
            byte[] csv = "Name,Age\nAda,0001\nBob,0002\nCid,0003\n"u8.ToArray();
            string path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            await File.WriteAllBytesAsync(path, csv, TestContext.Current.CancellationToken);
            try
            {
                using var headerReader = Excel.FromCsv(csv);
                CsvBoundColumnMap<Row> map = CsvHeaderBinder.Bind<Row>(
                    headerReader, new ExcelParserConfig(), TypeMapper<Row>.GetCsvInfo(), out _);

                using SafeFileHandle handle = File.OpenHandle(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                var source = new CsvChunkSource(handle, csv.Length);

                // Same chunk bounds as ParsesOnlyTheRecordsStartingInsideItsChunk, exercised through the
                // SafeFileHandle/RangedFileStream branch instead of the ReadOnlyMemory<byte> one.
                CsvChunkResult<Row> result = await CsvChunkWorker.ParseAsync(
                    source,
                    new CsvChunk(0, 18, 27),
                    confirmedStart: 18,
                    map,
                    TypeMapper<Row>.GetCsvInfo(),
                    CsvReaderOptions.Default,
                    new ExcelParserConfig(),
                    TestContext.Current.CancellationToken);

                Row row = Assert.Single(result.Models);
                Assert.Equal("Bob", row.Name);
                Assert.Equal(2, row.Age);
            }
            finally
            {
                File.Delete(path);
            }
        }
    }
}
