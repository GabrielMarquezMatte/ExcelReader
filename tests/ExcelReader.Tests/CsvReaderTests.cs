using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class CsvReaderTests
    {
        private static MemoryStream Csv(string content)
        {
            return new MemoryStream(Encoding.UTF8.GetBytes(content));
        }

        private static List<string[]> ReadAll(CsvReader reader)
        {
            var rows = new List<string[]>();
            using CsvReader.Enumerator e = reader.GetEnumerator();
            while (e.MoveNext())
            {
                var row = e.Current;
                var cells = new string[row.ColumnCount];
                for (int i = 0; i < row.ColumnCount; i++)
                {
                    cells[i] = row[i].GetString();
                }
                rows.Add(cells);
            }
            return rows;
        }

        private static async Task<List<string[]>> ReadAllAsync(CsvReader reader)
        {
            var rows = new List<string[]>();
            await using CsvReader.Enumerator e = await reader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);
            while (await e.MoveNextAsync())
            {
                // Row is a ref struct: extracted into a plain sync method so no ref-struct local
                // lives inside this async method's body (pre-C#13 async state machines disallow that).
                rows.Add(ToArray(e));
            }
            return rows;
        }

        private static string[] ToArray(CsvReader.Enumerator e)
        {
            var row = e.Current;
            var cells = new string[row.ColumnCount];
            for (int i = 0; i < row.ColumnCount; i++)
            {
                cells[i] = row[i].GetString();
            }
            return cells;
        }

        [Fact]
        public void SimpleRowsAreRead()
        {
            using var ms = Csv("a,b,c\n1,2,3\n");
            using var reader = Excel.FromCsv(ms);

            var rows = ReadAll(reader);

            Assert.Equal(2, rows.Count);
            Assert.Equal(["a", "b", "c"], rows[0]);
            Assert.Equal(["1", "2", "3"], rows[1]);
        }

        [Fact]
        public async Task SimpleRowsAreReadAsync()
        {
            using var ms = Csv("a,b,c\n1,2,3\n");
            using var reader = Excel.FromCsv(ms);

            var rows = await ReadAllAsync(reader);

            Assert.Equal(2, rows.Count);
            Assert.Equal(["a", "b", "c"], rows[0]);
            Assert.Equal(["1", "2", "3"], rows[1]);
        }

        [Fact]
        public void DelimiterInsideQuotesIsLiteral()
        {
            using var ms = Csv("""a,"b,c",d""");
            using var reader = Excel.FromCsv(ms);

            var rows = ReadAll(reader);

            Assert.Equal(["a", "b,c", "d"], Assert.Single(rows));
        }

        [Fact]
        public void EmbeddedNewlineInsideQuotesIsPreserved()
        {
            using var ms = Csv("\"line1\nline2\",b");
            using var reader = Excel.FromCsv(ms);

            var rows = ReadAll(reader);

            var row = Assert.Single(rows);
            Assert.Equal("line1\nline2", row[0]);
            Assert.Equal("b", row[1]);
        }

        [Fact]
        public void EscapedQuoteIsUnescaped()
        {
            using var ms = Csv("a,\"she said \"\"hi\"\"\",c");
            using var reader = Excel.FromCsv(ms);

            var rows = ReadAll(reader);

            Assert.Equal(["a", "she said \"hi\"", "c"], Assert.Single(rows));
        }

        [Fact]
        public void EmptyFieldsAreEmptyCellType()
        {
            using var ms = Csv("a,,c\n");
            using var reader = Excel.FromCsv(ms);
            using CsvReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            var row = e.Current;
            Assert.Equal(3, row.ColumnCount);
            Assert.Equal(CellType.Empty, row[1].Type);
            Assert.Equal(CellType.ExcelString, row[0].Type);
        }

        [Fact]
        public void TrailingDelimiterYieldsTrailingEmptyField()
        {
            using var ms = Csv("a,b,\n");
            using var reader = Excel.FromCsv(ms);

            var rows = ReadAll(reader);

            Assert.Equal(["a", "b", ""], Assert.Single(rows));
        }

        [Fact]
        public void EmptyQuotedFieldIsEmpty()
        {
            using var ms = Csv("""a,"",c""");
            using var reader = Excel.FromCsv(ms);
            using CsvReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            var row = e.Current;
            Assert.Equal(CellType.Empty, row[1].Type);
        }

        [Theory]
        [InlineData("a,b\r\n1,2\r\n")]
        [InlineData("a,b\n1,2\n")]
        [InlineData("a,b\r1,2\r")]
        public void AllLineTerminatorsAreSupported(string csv)
        {
            using var ms = Csv(csv);
            using var reader = Excel.FromCsv(ms);

            var rows = ReadAll(reader);

            Assert.Equal(2, rows.Count);
            Assert.Equal(["a", "b"], rows[0]);
            Assert.Equal(["1", "2"], rows[1]);
        }

        [Fact]
        public void MixedLineTerminatorsWithinOneFileAreSupported()
        {
            using var ms = Csv("a,b\r\nc,d\ne,f\r");
            using var reader = Excel.FromCsv(ms);

            var rows = ReadAll(reader);

            Assert.Equal(3, rows.Count);
            Assert.Equal(["a", "b"], rows[0]);
            Assert.Equal(["c", "d"], rows[1]);
            Assert.Equal(["e", "f"], rows[2]);
        }

        [Fact]
        public void NoTrailingNewlineStillYieldsFinalRecord()
        {
            using var ms = Csv("a,b,c");
            using var reader = Excel.FromCsv(ms);

            var rows = ReadAll(reader);

            Assert.Equal(["a", "b", "c"], Assert.Single(rows));
        }

        [Fact]
        public void TrailingNewlineDoesNotYieldPhantomRow()
        {
            using var ms = Csv("a,b,c\n");
            using var reader = Excel.FromCsv(ms);

            var rows = ReadAll(reader);

            Assert.Single(rows);
        }

        [Fact]
        public void EmptyFileYieldsNoRows()
        {
            using var ms = Csv("");
            using var reader = Excel.FromCsv(ms);

            var rows = ReadAll(reader);

            Assert.Empty(rows);
        }

        [Fact]
        public void BlankLineYieldsOneEmptyField()
        {
            using var ms = Csv("a,b\n\nc,d\n");
            using var reader = Excel.FromCsv(ms);

            var rows = ReadAll(reader);

            Assert.Equal(3, rows.Count);
            Assert.Equal([""], rows[1]);
        }

        [Fact]
        public void LargeFieldExceedingInitialBufferIsReadInFull()
        {
            string big = new('x', 200 * 1024); // larger than the 64KB initial scan buffer
            using var ms = Csv($"a,{big},c\n");
            using var reader = Excel.FromCsv(ms);

            var rows = ReadAll(reader);

            Assert.Equal(big, Assert.Single(rows)[1]);
        }

        [Fact]
        public void QuotedFieldSplitAcrossBufferBoundaryIsHandled()
        {
            // Force the quoted field's closing quote to fall right at/after typical buffer refills.
            string big = new('y', 130 * 1024);
            using var ms = Csv($"a,\"{big}\",c\n");
            using var reader = Excel.FromCsv(ms);

            var rows = ReadAll(reader);

            Assert.Equal(big, Assert.Single(rows)[1]);
        }

        [Fact]
        public async Task QuotedFieldSplitAcrossBufferBoundaryIsHandledAsync()
        {
            string big = new('z', 130 * 1024);
            using var ms = Csv($"a,\"{big}\",c\n");
            using var reader = Excel.FromCsv(ms);

            var rows = await ReadAllAsync(reader);

            Assert.Equal(big, Assert.Single(rows)[1]);
        }

        [Fact]
        public void FieldExceedingMaxCellBytesThrows()
        {
            // Must exceed the initial buffer so growth (and the limit check) actually happens.
            using var ms = Csv("a," + new string('x', 128 * 1024) + ",c\n");
            var options = new CsvReaderOptions { MaxCellBytes = 1024 };
            using var reader = Excel.FromCsv(ms, options: options);
            using CsvReader.Enumerator e = reader.GetEnumerator();

            ExcelLimitExceededException ex = Assert.Throws<ExcelLimitExceededException>(() => e.MoveNext());
            Assert.Equal(nameof(CsvReaderOptions.MaxCellBytes), ex.LimitName);
        }

        [Fact]
        public void SemicolonDelimiterIsRespected()
        {
            using var ms = Csv("a;b;c\n");
            var options = new CsvReaderOptions { Delimiter = (byte)';' };
            using var reader = Excel.FromCsv(ms, options: options);

            var rows = ReadAll(reader);

            Assert.Equal(["a", "b", "c"], Assert.Single(rows));
        }

        [Fact]
        public void TabDelimiterIsRespected()
        {
            using var ms = Csv("a\tb\tc\n");
            var options = new CsvReaderOptions { Delimiter = (byte)'\t' };
            using var reader = Excel.FromCsv(ms, options: options);

            var rows = ReadAll(reader);

            Assert.Equal(["a", "b", "c"], Assert.Single(rows));
        }

        [Fact]
        public void Utf8BomIsStripped()
        {
            byte[] bom = [0xEF, 0xBB, 0xBF];
            byte[] body = Encoding.UTF8.GetBytes("a,b\n");
            using var ms = new MemoryStream([.. bom, .. body]);
            using var reader = Excel.FromCsv(ms);

            var rows = ReadAll(reader);

            Assert.Equal(["a", "b"], Assert.Single(rows));
        }

        [Fact]
        public void MultibyteUtf8ContentIsPreserved()
        {
            using var ms = Csv("nome,cidade\nGabriel,São Paulo\n");
            using var reader = Excel.FromCsv(ms);

            var rows = ReadAll(reader);

            Assert.Equal("São Paulo", rows[1][1]);
        }

        [Fact]
        public void IsDate1904IsAlwaysFalse()
        {
            using var ms = Csv("a\n");
            using var reader = Excel.FromCsv(ms);

            Assert.False(reader.IsDate1904);
        }

        [Fact]
        public void DelimiterEqualToQuoteThrows()
        {
            using var ms = Csv("a,b\n");
            var options = new CsvReaderOptions { Delimiter = (byte)'"' };

            Assert.Throws<ArgumentException>(() => Excel.FromCsv(ms, options: options));
        }

        [Fact]
        public void SeekableStreamCanBeEnumeratedMoreThanOnce()
        {
            using var ms = Csv("a,b\n1,2\n");
            using var reader = Excel.FromCsv(ms);

            var first = ReadAll(reader);
            var second = ReadAll(reader);

            Assert.Equal(first.Count, second.Count);
            Assert.Equal(first[0], second[0]);
        }
    }
}
