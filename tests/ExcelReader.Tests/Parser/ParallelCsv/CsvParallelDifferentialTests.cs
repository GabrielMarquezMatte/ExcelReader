using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.ParallelCsv;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests.Parser.ParallelCsv
{
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
            foreach (Row row in ExcelParser.FromAttributes<Row>().Parse(reader.FirstSheet))
            {
                list.Add(Render(row));
            }
            return list;
        }

        private static async Task<List<string>> ParallelAsync(byte[] csv, int dop, int chunkSize)
        {
            var list = new List<string>();
            await foreach (Row row in ParallelCsvFactory.CreateWithChunkSize(
                csv.AsMemory(), dop, chunkSize, null, ExcelParser.FromAttributes<Row>(), CancellationToken.None))
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
                "Name,Age,Note\nAda,36,x\nBob,41,y\nCid,7,z\n",
                "Name,Age,Note\nAda,36,x\nBob,41,y",
                "Name,Age,Note\r\nAda,36,x\r\nBob,41,y\r\n",
                "Name,Age,Note\n\"multi\nline\",1,x\n\"another\none\",2,y\n",
                "Name,Age,Note\n\"say \"\"hi\"\"\",1,x\n\"and \"\"bye\"\"\",2,y\n",
                "Name,Age,Note\n\"a,b\",1,x\n\"c,d\",2,y\n",
                "Name,Age,Note\n,1,\n,2,\n",
                "Name,Age,Note\n\"" + new string('q', 400) + "\",1,x\n",
                "Name,Age,Note\nAda,36,x\n\nBob,41,y\n",
                "Name,Age,Note\rAda,36,x\rBob,41,y\r",
                "Name,Age,Note\nAda,36,x\r\nBob,41,y\rCid,7,z\n",
                "Name,Age,Note\nab\"c,1,x\nd\"e,2,y\n",
                "Name,Age,Note\n\"a\"b,1,x\n\"c\"d,2,\"e\nf\"\n",
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
        public async Task ReportsTheSequentialFailureRowAcrossQuotedNewlinesAndBlankLines()
        {
            byte[] csv = Encoding.UTF8.GetBytes("junk\nName,Age,Note\n\"a\nb\",1,x\n\nc,2,\"d\r\ne\"\r\n\n\"f\ng\",bad,y\nh,3,z\n");
            var options = new ExcelParserOptions { HeaderRow = 2, ThrowOnParseFailure = true };
            ExcelParser<Row> parser = ExcelParser.FromAttributes<Row>(options);
            using var reader = Excel.FromCsv(csv);
            int expected = Assert.Throws<ExcelParseException>(() => parser.Parse(reader.FirstSheet).ToList()).Row;

            for (int chunkSize = 1; chunkSize <= csv.Length; chunkSize++)
            {
                ExcelParseException parallel = await Assert.ThrowsAsync<ExcelParseException>(async () =>
                {
                    await foreach (Row _ in ParallelCsvFactory.CreateWithChunkSize(csv.AsMemory(), 3, chunkSize, null, parser, CancellationToken.None))
                    {
                    }
                });
                Assert.Equal(expected, parallel.Row);
            }
            Assert.Equal(7, expected);
        }

        [Fact]
        public async Task MatchesSequentialOnALargeGeneratedFile()
        {
            var sb = new StringBuilder("Name,Age,Note\n");
            for (int i = 0; i < 50_000; i++)
            {
                string name = i % 8 == 0 ? $"\"multi\nline{i}\"" : $"name{i}";
                sb.Append(CultureInfo.InvariantCulture, $"{name},{i},note{i}\n");
            }
            byte[] csv = Encoding.UTF8.GetBytes(sb.ToString());
            List<string> expected = Sequential(csv);

            List<string> actual = await ParallelAsync(csv, dop: 8, chunkSize: 4096);

            Assert.Equal(expected, actual);
        }

        private struct StructRow
        {
            public string? Name { get; set; }
            public int Age { get; set; }
            public string? Note { get; set; }
        }

        private static byte[] MultiChunkCsv()
        {
            var sb = new StringBuilder("Name,Age,Note\n");
            for (int i = 0; i < 3_000; i++)
            {
                string name = i % 5 == 0 ? $"\"multi\nline{i}\"" : $"name{i}";
                sb.Append(CultureInfo.InvariantCulture, $"{name},{i},note{i}\n");
                if (i % 97 == 0)
                {
                    sb.Append('\n');
                }
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        private static async Task<List<string>> ParallelStructAsync(byte[] csv, int dop, int chunkSize)
        {
            var list = new List<string>();
            await foreach (StructRow row in ParallelCsvFactory.CreateWithChunkSize(
                csv.AsMemory(), dop, chunkSize, null, ExcelParser.FromAttributes<StructRow>(), CancellationToken.None))
            {
                list.Add($"{row.Name}{row.Age}{row.Note}");
            }
            return list;
        }

        [Fact]
        public async Task StructAndClassModelsMatchSequentialAcrossManyChunks()
        {
            byte[] csv = MultiChunkCsv();
            List<string> expected = Sequential(csv);
            using var reader = Excel.FromCsv(csv);
            var expectedStructs = new List<string>();
            foreach (StructRow row in ExcelParser.FromAttributes<StructRow>().Parse(reader.FirstSheet))
            {
                expectedStructs.Add($"{row.Name}{row.Age}{row.Note}");
            }
            Assert.Equal(3_000, expected.Count);
            Assert.Equal(expected, expectedStructs);

            foreach (int dop in new[] { 2, 4, 8 })
            {
                foreach (int chunkSize in new[] { 64, 1_000, 4_096, 16_384 })
                {
                    Assert.Equal(expected, await ParallelAsync(csv, dop, chunkSize));
                    Assert.Equal(expected, await ParallelStructAsync(csv, dop, chunkSize));
                }
            }
        }
    }
}
