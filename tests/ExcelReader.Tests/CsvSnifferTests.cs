using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class CsvSnifferTests
    {
        private static byte[] Bytes(string s)
        {
            return Encoding.UTF8.GetBytes(s);
        }

        private static byte[] RepeatedRows(string row, int count)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < count; i++)
            {
                sb.Append(row);
            }
            return Bytes(sb.ToString());
        }

        [Theory]
        [InlineData((byte)',')]
        [InlineData((byte)';')]
        [InlineData((byte)'\t')]
        [InlineData((byte)'|')]
        public void DetectsDelimiter(byte delimiter)
        {
            char d = (char)delimiter;
            byte[] sample = RepeatedRows($"a{d}b{d}c\n", 10);
            CsvDialect dialect = CsvSniffer.Detect(sample);
            Assert.Equal(delimiter, dialect.Delimiter);
            Assert.Equal((byte)'"', dialect.Quote);
        }

        [Fact]
        public void DelimiterInsideQuotesIsNotCounted()
        {
            // Real delimiter is ';'. The number of commas embedded in the quoted field varies line
            // to line, so a naive ',' candidate (which, unquoted, counts them as real delimiters)
            // scores inconsistently and must lose to ';', even though ',' comes first in candidate
            // order.
            var sb = new StringBuilder();
            string[] rows = ["a;\"b,c\";d\n", "a;\"b,c,e\";d\n", "a;\"b\";d\n"];
            for (int i = 0; i < 6; i++)
            {
                sb.Append(rows[i % rows.Length]);
            }
            CsvDialect dialect = CsvSniffer.Detect(Encoding.UTF8.GetBytes(sb.ToString()));
            Assert.Equal((byte)';', dialect.Delimiter);
        }

        [Fact]
        public void SingleColumnFileReturnsDefault()
        {
            byte[] sample = RepeatedRows("onlyonecolumn\n", 6);
            CsvDialect dialect = CsvSniffer.Detect(sample);
            Assert.Equal(CsvDialect.Default, dialect);
        }

        [Fact]
        public void EmptySampleReturnsDefault()
        {
            CsvDialect dialect = CsvSniffer.Detect(ReadOnlySpan<byte>.Empty);
            Assert.Equal(CsvDialect.Default, dialect);
        }

        [Theory]
        [InlineData(6)]
        [InlineData(12)]
        [InlineData(20)]
        [InlineData(27)]
        public void TruncatedLastLineDoesNotChangeResult(int cutOffset)
        {
            byte[] full = RepeatedRows("field1,field2,field3\n", 10);
            byte[] truncated = full[..(full.Length - cutOffset)];
            CsvDialect dialect = CsvSniffer.Detect(truncated);
            Assert.Equal((byte)',', dialect.Delimiter);
            Assert.Equal((byte)'"', dialect.Quote);
        }

        [Fact]
        public void TieBreaksByCandidateOrder()
        {
            // Every line has exactly one ',' and one ';' — both delimiters score a perfectly
            // consistent (variance-zero) two-field split. ',' must win because it comes first in
            // CsvSnifferOptions.Default.CandidateDelimiters.
            byte[] sample = RepeatedRows("a,b;c\n", 6);
            CsvDialect dialect = CsvSniffer.Detect(sample);
            Assert.Equal((byte)',', dialect.Delimiter);
        }

        [Fact]
        public void DetectsUtf8Bom()
        {
            byte[] bom = [0xEF, 0xBB, 0xBF];
            byte[] body = RepeatedRows("a,b,c\n", 6);
            byte[] sample = [.. bom, .. body];
            CsvDialect dialect = CsvSniffer.Detect(sample);
            Assert.True(dialect.HasByteOrderMark);
            Assert.Equal(Encoding.UTF8, dialect.Encoding);
            Assert.Equal((byte)',', dialect.Delimiter);
        }

        [Fact]
        public void DetectsUtf16LeBom()
        {
            byte[] bom = [0xFF, 0xFE];
            byte[] body = RepeatedRows("a,b,c\n", 6);
            byte[] sample = [.. bom, .. body];
            CsvDialect dialect = CsvSniffer.Detect(sample);
            Assert.True(dialect.HasByteOrderMark);
            Assert.Equal(Encoding.Unicode, dialect.Encoding);
        }

        [Fact]
        public void SingleQuoteDialectIsDetected()
        {
            // Delimiter is ';'. Each line embeds a different number of ';' inside the single-quoted
            // field, so only the quote="'" interpretation gives a consistent field count per line —
            // quote='"' (never present) counts every embedded ';' as a real delimiter, which varies
            // line to line and must lose.
            var sb = new StringBuilder();
            string[] rows = ["a;'x;y';b\n", "a;'x;y;z';b\n", "a;'x';b\n"];
            for (int i = 0; i < 6; i++)
            {
                sb.Append(rows[i % rows.Length]);
            }
            CsvDialect dialect = CsvSniffer.Detect(Bytes(sb.ToString()));
            Assert.Equal((byte)';', dialect.Delimiter);
            Assert.Equal((byte)'\'', dialect.Quote);
        }

        [Fact]
        public void NonSeekableStreamThrowsArgumentException()
        {
            using var stream = new NonSeekableStream(RepeatedRows("a,b,c\n", 6));
            Assert.Throws<ArgumentException>(() => Excel.SniffCsvDialect(stream));
        }

        [Fact]
        public void SniffedDialectRoundTripsThroughReader()
        {
            byte[] data = Bytes("nome;idade\nAna;30\nBia;25\n");
            CsvDialect dialect = CsvSniffer.Detect(data);
            CsvReaderOptions options = CsvReaderOptions.Default.WithDialect(dialect);
            using var reader = Excel.FromCsv(data, options);
            using CsvReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("nome", e.Current[0].GetString());
            Assert.Equal("idade", e.Current[1].GetString());
            Assert.True(e.MoveNext());
            Assert.Equal("Ana", e.Current[0].GetString());
            Assert.Equal("30", e.Current[1].GetString());
            Assert.True(e.MoveNext());
            Assert.Equal("Bia", e.Current[0].GetString());
            Assert.Equal("25", e.Current[1].GetString());
            Assert.False(e.MoveNext());
        }

        // Deliberately not cryptographically secure — CA5394 doesn't apply: a seeded, reproducible
        // PRNG is exactly what makes a sniffer failure on random input reproducible.
        [SuppressMessage("Security", "CA5394:Do not use insecure randomness",
            Justification = "Fuzzing the sniffer needs a reproducible seeded PRNG, not cryptographic randomness.")]
        [Theory]
        [InlineData(1)]
        [InlineData(2)]
        [InlineData(3)]
        [InlineData(4)]
        [InlineData(5)]
        [InlineData(6)]
        [InlineData(7)]
        [InlineData(8)]
        public void RandomizedSnifferNeverThrows(int seed)
        {
            var random = new Random(seed);
            byte[] data = new byte[512];
            random.NextBytes(data);
            CsvDialect dialect = CsvSniffer.Detect(data);
            Assert.Contains(dialect.Delimiter, CsvSnifferOptions.Default.CandidateDelimiters);
        }
    }
}
