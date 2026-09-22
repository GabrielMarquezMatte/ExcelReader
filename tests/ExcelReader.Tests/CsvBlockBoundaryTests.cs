using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class CsvBlockBoundaryTests
    {
        public static TheoryData<int> Offsets()
        {
            var data = new TheoryData<int>();
            for (int pad = 0; pad <= 130; pad++)
            {
                data.Add(pad);
            }
            return data;
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

        private static void AssertBothSources(string csv, params string[][] expected)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(csv);
            using (var fromMemory = Excel.FromCsv(bytes))
            {
                Assert.Equal(expected, ReadAll(fromMemory));
            }
            using var ms = new MemoryStream(bytes);
            using var fromStream = Excel.FromCsv(ms);
            Assert.Equal(expected, ReadAll(fromStream));
        }

        [Theory]
        [MemberData(nameof(Offsets))]
        public void QuotedFieldWithEscapesDelimitersAndNewlinesParsesAtEveryBlockOffset(int pad)
        {
            string lead = new('x', pad);
            string csv = lead + ",\"a\"\"b,c\r\nd\"\"\",\"\",e\r\n\"f\"\r\ng";

            AssertBothSources(csv, [lead, "a\"b,c\r\nd\"", "", "e"], ["f"], ["g"]);
        }

        [Theory]
        [MemberData(nameof(Offsets))]
        public void CrLfIsOneTerminatorAtEveryBlockOffset(int pad)
        {
            string lead = new('x', pad);

            AssertBothSources(lead + "\r\n\"q\"\r\n\r\nz", [lead], ["q"], [""], ["z"]);
        }

        [Fact]
        public void EscapedQuotesSpanningSeveralBlocksAreUnescaped()
        {
            var body = new StringBuilder();
            for (int i = 0; i < 40; i++)
            {
                body.Append("ab\"\"cd,\n");
            }
            string escaped = body.ToString();
            string expected = escaped.Replace("\"\"", "\"", StringComparison.Ordinal);

            AssertBothSources("k,\"" + escaped + "\",tail\nnext", ["k", expected, "tail"], ["next"]);
        }

        [Fact]
        public void EscapedFieldIsReportedAsMaterializedString()
        {
            using var reader = Excel.FromCsv(Encoding.UTF8.GetBytes("\"a\"\"b\",\"\""));
            using CsvReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(CellType.ExcelString, e.Current[0].Type);
            Assert.Equal("a\"b", e.Current[0].GetString());
            Assert.Equal(CellType.Empty, e.Current[1].Type);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(62)]
        [InlineData(63)]
        [InlineData(64)]
        [InlineData(100)]
        public void MalformedQuotesAfterCleanPrefixKeepLenientSemantics(int pad)
        {
            string lead = new('x', pad);

            AssertBothSources(lead + ",b\"c,\"d\"e\nnext,\"n\"", [lead, "b\"c", "de"], ["next", "n"]);
        }

        [Theory]
        [InlineData(0)]
        [InlineData(63)]
        [InlineData(64)]
        public void UnclosedQuoteRunsToEndOfInput(int pad)
        {
            string lead = new('x', pad);

            AssertBothSources(lead + ",\"b,c\nd", [lead, "b,c\nd"]);
        }

        [Fact]
        public void RecordsOfVaryingWidthKeepTheirFieldsAndStartOffsets()
        {
            var csv = new StringBuilder();
            var expectedRows = new List<string[]>();
            var expectedStarts = new List<long>();
            for (int r = 0; r < 600; r++)
            {
                int width = r % 50 == 0 ? 300 : (r % 7) + 1;
                string[] row = new string[width];
                for (int c = 0; c < width; c++)
                {
                    row[c] = (r % 11 == 0 && c == 0) ? $"q\"{r}" : $"{r}.{c}";
                }
                expectedStarts.Add(csv.Length);
                expectedRows.Add(row);
                csv.Append(string.Join(',', row.Select(QuoteIfNeeded))).Append(r % 3 == 0 ? "\r\n" : "\n");
            }
            byte[] bytes = Encoding.UTF8.GetBytes(csv.ToString());

            var rows = new List<string[]>();
            var starts = new List<long>();
            using var e = new CsvReader.Enumerator(bytes, CsvReaderOptions.Default);
            while (e.MoveNext())
            {
                starts.Add(e.CurrentRecordStart);
                var fields = new string[e.FieldCount];
                for (int i = 0; i < fields.Length; i++)
                {
                    fields[i] = e.FieldAt(i).GetString();
                }
                rows.Add(fields);
            }

            Assert.Equal(expectedRows, rows);
            Assert.Equal(expectedStarts, starts);
        }

        private static string QuoteIfNeeded(string field)
        {
            if (!field.Contains('"', StringComparison.Ordinal))
            {
                return field;
            }
            return "\"" + field.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";
        }

        [Fact]
        public void QuoteEqualToDelimiterIsRejectedByBothSourcesAlike()
        {
            var options = CsvReaderOptions.Default with { Delimiter = (byte)',', Quote = (byte)',' };
            byte[] bytes = Encoding.UTF8.GetBytes("a,,b,c\n,d,\ne");

            Assert.Throws<ArgumentException>(() => Excel.FromCsv(bytes, options));

            using var ms = new MemoryStream(bytes);
            Assert.Throws<ArgumentException>(() => Excel.FromCsv(ms, options: options));
        }
    }
}
