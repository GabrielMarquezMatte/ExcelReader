using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    // The parallel path's only meaningful oracle is the sequential parser. Every test here compares
    // against it rather than against hand-written expectations.
    public class CsvParallelDifferentialTests
    {
        private sealed class Row
        {
            public string? Name { get; set; }
            public int Age { get; set; }
            public string? Note { get; set; }
        }

        private static List<string> Sequential(byte[] csv)
        {
            using var reader = Excel.FromCsv(csv);
            var list = new List<string>();
            foreach (Row row in new ExcelParser<Row>().Parse(reader))
            {
                list.Add(Render(row));
            }
            return list;
        }

        private static async Task<List<string>> ParallelAsync(byte[] csv, int dop, int chunkSize)
        {
            var list = new List<string>();
            await foreach (Row row in ParallelCsvFactory.CreateWithChunkSize<Row>(
                csv.AsMemory(), dop, chunkSize, null, null, CancellationToken.None))
            {
                list.Add(Render(row));
            }
            return list;
        }

        private static string Render(Row r)
        {
            return $"{r.Name}{r.Age}{r.Note}";
        }

        public static TheoryData<string> Corpus()
        {
            var data = new TheoryData<string>
            {
                // Plain.
                "Name,Age,Note\nAda,36,x\nBob,41,y\nCid,7,z\n",
                // No trailing newline.
                "Name,Age,Note\nAda,36,x\nBob,41,y",
                // CRLF throughout.
                "Name,Age,Note\r\nAda,36,x\r\nBob,41,y\r\n",
                // Quoted field carrying a newline.
                "Name,Age,Note\n\"multi\nline\",1,x\n\"another\none\",2,y\n",
                // Doubled quotes.
                "Name,Age,Note\n\"say \"\"hi\"\"\",1,x\n\"and \"\"bye\"\"\",2,y\n",
                // Quoted field carrying the delimiter.
                "Name,Age,Note\n\"a,b\",1,x\n\"c,d\",2,y\n",
                // Empty fields.
                "Name,Age,Note\n,1,\n,2,\n",
                // A single record that is one big quoted field.
                "Name,Age,Note\n\"" + new string('q', 400) + "\",1,x\n",
                // Blank line in the middle.
                "Name,Age,Note\nAda,36,x\n\nBob,41,y\n",
                // Lone CR terminators.
                "Name,Age,Note\rAda,36,x\rBob,41,y\r",
            };
            return data;
        }

        [Theory]
        [MemberData(nameof(Corpus))]
        public async Task MatchesSequentialAcrossEveryDegreeAndChunkSize(string text)
        {
            byte[] csv = Encoding.UTF8.GetBytes(text);
            List<string> expected = Sequential(csv);

            foreach (int dop in new[] { 2, 3, 4, 8 })
            {
                foreach (int chunkSize in new[] { 1, 2, 3, 7, 16, 64 })
                {
                    List<string> actual = await ParallelAsync(csv, dop, chunkSize);
                    Assert.Equal(expected, actual);
                }
            }
        }

        [Fact]
        public async Task SweepsAChunkBoundaryAcrossEveryByteOffsetOfAQuotedNewlineFixture()
        {
            // A chunk size of 1 puts a boundary between every pair of bytes, so this single test
            // exercises every possible misalignment against a file whose quoted fields contain the
            // one byte a naive splitter would trust.
            byte[] csv = Encoding.UTF8.GetBytes(
                "Name,Age,Note\n\"a\nb\",1,\"c\nd\"\n\"e\nf\",2,\"g\nh\"\n\"i\nj\",3,\"k\nl\"\n");
            List<string> expected = Sequential(csv);

            for (int chunkSize = 1; chunkSize <= csv.Length; chunkSize++)
            {
                List<string> actual = await ParallelAsync(csv, dop: 4, chunkSize);
                Assert.Equal(expected, actual);
            }
        }

        [Fact]
        public async Task MatchesSequentialOnALargeGeneratedFile()
        {
            var sb = new StringBuilder("Name,Age,Note\n");
            for (int i = 0; i < 50_000; i++)
            {
                // Every eighth row carries an embedded newline, so boundary corrections fire often.
                string name = i % 8 == 0 ? $"\"multi\nline{i}\"" : $"name{i}";
                sb.Append(CultureInfo.InvariantCulture, $"{name},{i},note{i}\n");
            }
            byte[] csv = Encoding.UTF8.GetBytes(sb.ToString());
            List<string> expected = Sequential(csv);

            List<string> actual = await ParallelAsync(csv, dop: 8, chunkSize: 4096);

            Assert.Equal(expected, actual);
        }
    }
}
