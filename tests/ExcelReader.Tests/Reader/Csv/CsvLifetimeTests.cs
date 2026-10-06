using System.Globalization;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Tests.Reader.Csv
{
    public sealed class CsvLifetimeTests : IDisposable
    {
        private const int Rows = 4000;

        private readonly List<string> _paths = [];

        public void Dispose()
        {
            foreach (string path in _paths)
            {
                File.Delete(path);
            }
        }

        private static byte[] BuildCsv(Encoding? encoding = null)
        {
            StringBuilder text = new();
            for (int r = 0; r < Rows; r++)
            {
                text.Append("name-").Append(r.ToString(CultureInfo.InvariantCulture))
                    .Append(',').Append((r * 7).ToString(CultureInfo.InvariantCulture))
                    .Append(",\"quoted, ").Append((r % 13).ToString(CultureInfo.InvariantCulture)).Append("\"\n");
            }
            return (encoding ?? Encoding.UTF8).GetBytes(text.ToString());
        }

        private static List<string> Drain(CsvReader.Enumerator e)
        {
            List<string> lines = [];
            while (e.MoveNext())
            {
                Row row = e.Current;
                lines.Add(row[0].GetString() + "|" + row[1].GetString() + "|" + row[2].GetString());
            }
            return lines;
        }

        private static List<string> Expected()
        {
            using CsvReader reference = Excel.FromCsv(BuildCsv());
            using CsvReader.Enumerator all = reference.FirstSheet.GetEnumerator();
            return Drain(all);
        }

        [Fact]
        public void Two_Enumerators_Over_One_Seekable_Stream_Do_Not_Share_A_Position()
        {
            List<string> expected = Expected();
            using CsvReader reader = Excel.FromCsv(new TrickleStream(BuildCsv()), leaveOpen: false);
            using CsvReader.Enumerator first = reader.FirstSheet.GetEnumerator();
            using CsvReader.Enumerator second = reader.FirstSheet.GetEnumerator();

            List<string> a = [];
            List<string> b = [];
            bool moreA = true;
            bool moreB = true;
            while (moreA || moreB)
            {
                moreA = moreA && first.MoveNext();
                if (moreA)
                {
                    a.Add(first.Current[0].GetString() + "|" + first.Current[1].GetString() + "|" + first.Current[2].GetString());
                }
                moreB = moreB && second.MoveNext();
                if (moreB)
                {
                    b.Add(second.Current[0].GetString() + "|" + second.Current[1].GetString() + "|" + second.Current[2].GetString());
                }
            }

            Assert.Equal(expected, a, StringComparer.Ordinal);
            Assert.Equal(expected, b, StringComparer.Ordinal);
        }

        [Theory]
        [InlineData("memory")]
        [InlineData("stream")]
        [InlineData("path")]
        public void Enumerators_On_Different_Threads_Each_Read_Every_Row(string source)
        {
            List<string> expected = Expected();
            byte[] bytes = BuildCsv();
            using CsvReader reader = Open(source, bytes);

            const int Threads = 6;
            CsvReader.Enumerator[] enumerators = new CsvReader.Enumerator[Threads];
            for (int t = 0; t < Threads; t++)
            {
                enumerators[t] = reader.FirstSheet.GetEnumerator();
            }
            List<string>[] actual = new List<string>[Threads];
            Parallel.For(0, Threads, new ParallelOptions { MaxDegreeOfParallelism = Threads }, t =>
            {
                using CsvReader.Enumerator e = enumerators[t];
                actual[t] = Drain(e);
            });

            Assert.All(actual, lines => Assert.Equal(expected, lines, StringComparer.Ordinal));
        }

        private CsvReader Open(string source, byte[] bytes)
        {
            if (string.Equals(source, "memory", StringComparison.Ordinal))
            {
                return Excel.FromCsv(bytes);
            }
            if (string.Equals(source, "stream", StringComparison.Ordinal))
            {
                return Excel.FromCsv(new TrickleStream(bytes), leaveOpen: false);
            }
            string path = Path.Combine(Path.GetTempPath(), "excelreader-csv-" + Guid.NewGuid().ToString("N") + ".csv");
            File.WriteAllBytes(path, bytes);
            _paths.Add(path);
            return Excel.FromCsvFile(path);
        }

        [Fact]
        public void A_Transcoded_Stream_Serves_Two_Enumerators()
        {
            List<string> expected = Expected();
            using CsvReader reader = Excel.FromCsv(
                new TrickleStream(BuildCsv(Encoding.Unicode)), leaveOpen: false,
                new CsvReaderOptions { Encoding = Encoding.Unicode, DetectEncodingFromByteOrderMark = false });
            using CsvReader.Enumerator first = reader.FirstSheet.GetEnumerator();
            using CsvReader.Enumerator second = reader.FirstSheet.GetEnumerator();
            Assert.Equal(expected, Drain(first), StringComparer.Ordinal);
            Assert.Equal(expected, Drain(second), StringComparer.Ordinal);
        }

        [Fact]
        public void A_Stream_Already_Advanced_By_The_Caller_Starts_Every_Enumerator_There()
        {
            List<string> expected = Expected();
            byte[] csv = BuildCsv();
            byte[] prefix = Encoding.UTF8.GetBytes("this line is not part of the data\n");
            TrickleStream stream = new([.. prefix, .. csv]);
            stream.Position = prefix.Length;

            using CsvReader reader = Excel.FromCsv(stream, leaveOpen: false);
            using CsvReader.Enumerator first = reader.FirstSheet.GetEnumerator();
            using CsvReader.Enumerator second = reader.FirstSheet.GetEnumerator();
            Assert.Equal(expected, Drain(first), StringComparer.Ordinal);
            Assert.Equal(expected, Drain(second), StringComparer.Ordinal);
        }

        [Fact]
        public void A_NonSeekable_Stream_Serves_Exactly_One_Enumerator()
        {
            using NonSeekableStream stream = new(BuildCsv());
            using CsvReader reader = Excel.FromCsv(stream, leaveOpen: true);

            int granted = 0;
            int refused = 0;
            CsvReader.Enumerator?[] taken = new CsvReader.Enumerator?[8];
            Parallel.For(0, taken.Length, new ParallelOptions { MaxDegreeOfParallelism = taken.Length }, t =>
            {
                try
                {
                    taken[t] = reader.FirstSheet.GetEnumerator();
                    Interlocked.Increment(ref granted);
                }
                catch (InvalidOperationException)
                {
                    Interlocked.Increment(ref refused);
                }
            });

            Assert.Equal(1, granted);
            Assert.Equal(taken.Length - 1, refused);
            using CsvReader.Enumerator only = taken.Single(e => e is not null)!;
            Assert.Equal(Expected(), Drain(only), StringComparer.Ordinal);
        }

        [Fact]
        public void An_Enumerator_Outlives_Its_Reader()
        {
            List<string> expected = Expected();
            TrickleStream stream = new(BuildCsv());
            CsvReader reader = Excel.FromCsv(stream, leaveOpen: false);
            CsvReader.Enumerator e = reader.FirstSheet.GetEnumerator();
            Assert.True(e.MoveNext());

            reader.Dispose();
            Assert.True(stream.CanRead);
            Assert.Equal(expected.Skip(1), Drain(e), StringComparer.Ordinal);

            e.Dispose();
            Assert.False(stream.CanRead);
        }

        [Fact]
        public void GetEnumerator_After_Dispose_Throws()
        {
            CsvReader reader = Excel.FromCsv(new TrickleStream(BuildCsv()), leaveOpen: false);
            CsvSheet sheet = reader.FirstSheet;
            reader.Dispose();
            Assert.Throws<ObjectDisposedException>(() => reader.Sheets);
            Assert.Throws<ObjectDisposedException>(() => sheet.GetEnumerator());
            Assert.Throws<ObjectDisposedException>(() => sheet.GetAsyncEnumerator(TestContext.Current.CancellationToken));
        }

        [Fact]
        public void A_Transcoded_File_Offers_No_Raw_Chunk_Source()
        {
            string path = Path.Combine(Path.GetTempPath(), "excelreader-csv-" + Guid.NewGuid().ToString("N") + ".csv");
            File.WriteAllBytes(path, BuildCsv(Encoding.Unicode));
            _paths.Add(path);

            using CsvReader utf16 = Excel.FromCsvFile(path, new CsvReaderOptions { Encoding = Encoding.Unicode, DetectEncodingFromByteOrderMark = false });
            Assert.False(utf16.TryGetChunkSource(out _));

            using CsvReader utf8 = Excel.FromCsvFile(path);
            Assert.True(utf8.TryGetChunkSource(out _));
        }

        [Fact]
        public void A_Borrowed_Stream_Is_Left_Open()
        {
            TrickleStream stream = new(BuildCsv());
            using (CsvReader reader = Excel.FromCsv(stream, leaveOpen: true))
            {
                using CsvReader.Enumerator e = reader.FirstSheet.GetEnumerator();
                Assert.True(e.MoveNext());
            }
            Assert.True(stream.CanRead);
        }
    }
}
