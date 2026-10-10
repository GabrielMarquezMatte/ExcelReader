using System.Globalization;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Schema;

namespace ExcelReader.Tests.Reader.Csv
{
    public sealed class CsvReplayTests
    {
        private static byte[] BuildCsv(int rows, Encoding? encoding = null)
        {
            StringBuilder text = new("name,value\n");
            for (int r = 0; r < rows; r++)
            {
                text.Append("name-").Append(r.ToString(CultureInfo.InvariantCulture))
                    .Append(',').Append((r * 7).ToString(CultureInfo.InvariantCulture)).Append('\n');
            }
            return (encoding ?? Encoding.UTF8).GetBytes(text.ToString());
        }

        private static List<string> Drain(CsvReader reader)
        {
            List<string> lines = [];
            using CsvReader.Enumerator e = reader.FirstSheet.GetEnumerator();
            while (e.MoveNext())
            {
                Row row = e.Current;
                lines.Add(row[0].GetString() + "|" + row[1].GetString());
            }
            return lines;
        }

        private static List<string> Expected(int rows)
        {
            using CsvReader reference = Excel.FromCsv(BuildCsv(rows));
            return Drain(reference);
        }

        [Fact]
        public void A_Schema_Inferred_From_A_NonSeekable_Stream_Leaves_Every_Row_To_Read()
        {
            const int rows = 20_000;
            using NonSeekableStream stream = new(BuildCsv(rows));
            using CsvReader reader = Excel.FromCsv(stream);

            ExcelColumnSchema[] schema = Excel.InferSchema(reader.FirstSheet, headerRow: 1, sampleSize: 100, parseText: true);

            Assert.Equal(["name", "value"], schema.Select(c => c.Name), StringComparer.Ordinal);
            Assert.Equal(ExcelColumnType.Int64Column, schema[1].Type);
            Assert.Equal(Expected(rows), Drain(reader), StringComparer.Ordinal);
        }

        [Fact]
        public void A_Small_NonSeekable_Stream_Can_Be_Read_Any_Number_Of_Times()
        {
            const int rows = 500;
            List<string> expected = Expected(rows);
            using NonSeekableStream stream = new(BuildCsv(rows));
            using CsvReader reader = Excel.FromCsv(stream);

            Assert.Equal(expected, Drain(reader), StringComparer.Ordinal);
            Assert.Equal(expected, Drain(reader), StringComparer.Ordinal);
            Assert.Equal(expected, Drain(reader), StringComparer.Ordinal);
        }

        [Fact]
        public void A_Transcoded_NonSeekable_Stream_Is_Replayed_Too()
        {
            const int rows = 20_000;
            using NonSeekableStream stream = new(BuildCsv(rows, Encoding.Unicode));
            using CsvReader reader = Excel.FromCsv(
                stream, leaveOpen: false,
                new CsvReaderOptions { Encoding = Encoding.Unicode, DetectEncodingFromByteOrderMark = false });

            Excel.InferSchema(reader.FirstSheet);

            Assert.Equal(Expected(rows), Drain(reader), StringComparer.Ordinal);
        }

        [Fact]
        public void A_Second_Enumeration_Is_Refused_While_The_First_Is_Open()
        {
            using NonSeekableStream stream = new(BuildCsv(10));
            using CsvReader reader = Excel.FromCsv(stream);
            using CsvReader.Enumerator first = reader.FirstSheet.GetEnumerator();

            Assert.Throws<InvalidOperationException>(reader.FirstSheet.GetEnumerator);
        }

        [Fact]
        public void A_Stream_Read_Past_The_Replay_Limit_Cannot_Be_Read_Again()
        {
            byte[] line = Encoding.UTF8.GetBytes(new string('x', 1021) + ",1\n");
            byte[] bytes = new byte[CsvReader.ReplayLimit + (line.Length * 4)];
            for (int offset = 0; offset < bytes.Length; offset += line.Length)
            {
                line.CopyTo(bytes, offset);
            }
            using NonSeekableStream stream = new(bytes);
            using CsvReader reader = Excel.FromCsv(stream);

            Assert.Equal(bytes.Length / line.Length, Drain(reader).Count);

            InvalidOperationException error = Assert.Throws<InvalidOperationException>(reader.FirstSheet.GetEnumerator);
            Assert.Contains("cannot be enumerated again", error.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void A_Pass_Replays_What_Was_Read_And_Then_Continues()
        {
            byte[] bytes = Enumerable.Range(0, 100).Select(i => (byte)i).ToArray();
            using NonSeekableStream inner = new(bytes);
            using ReplayStream replay = new(inner, leaveOpen: true, limit: 64);

            Assert.True(replay.TryBeginPass());
            byte[] head = new byte[40];
            replay.ReadExactly(head);
            replay.Dispose();

            Assert.True(replay.TryBeginPass());
            byte[] all = new byte[100];
            replay.ReadExactly(all);
            Assert.Equal(bytes, all);
            Assert.Equal(0, replay.Read(new byte[1]));
            replay.Dispose();

            Assert.False(replay.TryBeginPass());
        }

        [Fact]
        public void Releasing_Closes_The_Inner_Stream_Only_When_It_Is_Owned()
        {
            using TrackingStream owned = new();
            using TrackingStream borrowed = new();
            ReplayStream a = new(owned, leaveOpen: false, limit: 8);
            ReplayStream b = new(borrowed, leaveOpen: true, limit: 8);

            a.Dispose();
            Assert.False(owned.Disposed);
            a.Release();
            b.Release();

            Assert.True(owned.Disposed);
            Assert.False(borrowed.Disposed);
        }

        private sealed class TrackingStream : MemoryStream
        {
            internal bool Disposed { get; private set; }

            protected override void Dispose(bool disposing)
            {
                Disposed = true;
                base.Dispose(disposing);
            }
        }
    }
}
