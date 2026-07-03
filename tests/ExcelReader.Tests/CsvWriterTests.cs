using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    public class CsvWriterTests
    {
        private static string Write(Action<CsvWriter> build, CsvWriterOptions? options = null)
        {
            var ms = new MemoryStream();
            using (CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true, options: options))
            {
                build(writer);
            }
            return Encoding.UTF8.GetString(ms.ToArray());
        }

        [Fact]
        public void SimpleRowIsWritten()
        {
            string csv = Write(w =>
            {
                using CsvRowWriter row = w.StartRow();
                row.Write("a");
                row.Write("b");
                row.Write("c");
            });

            Assert.Equal("a,b,c\r\n", csv);
        }

        [Fact]
        public void MultipleRowsAreWritten()
        {
            string csv = Write(w =>
            {
                using (CsvRowWriter row = w.StartRow())
                {
                    row.Write("a");
                    row.Write("b");
                }
                using (CsvRowWriter row = w.StartRow())
                {
                    row.Write("1");
                    row.Write("2");
                }
            });

            Assert.Equal("a,b\r\n1,2\r\n", csv);
        }

        [Fact]
        public void FieldContainingDelimiterIsQuoted()
        {
            string csv = Write(w =>
            {
                using CsvRowWriter row = w.StartRow();
                row.Write("a,b");
                row.Write("c");
            });

            Assert.Equal("\"a,b\",c\r\n", csv);
        }

        [Fact]
        public void FieldContainingQuoteIsEscaped()
        {
            string csv = Write(w =>
            {
                using CsvRowWriter row = w.StartRow();
                row.Write("she said \"hi\"");
            });

            Assert.Equal("\"she said \"\"hi\"\"\"\r\n", csv);
        }

        [Fact]
        public void FieldContainingNewlineIsQuoted()
        {
            string csv = Write(w =>
            {
                using CsvRowWriter row = w.StartRow();
                row.Write("line1\nline2");
                row.Write("b");
            });

            Assert.Equal("\"line1\nline2\",b\r\n", csv);
        }

        [Fact]
        public void NullStringWritesEmptyField()
        {
            string csv = Write(w =>
            {
                using CsvRowWriter row = w.StartRow();
                row.Write("a");
                row.Write((string?)null);
                row.Write("c");
            });

            Assert.Equal("a,,c\r\n", csv);
        }

        [Fact]
        public void SkipWritesEmptyFields()
        {
            string csv = Write(w =>
            {
                using CsvRowWriter row = w.StartRow();
                row.Write("a");
                row.Skip(2);
                row.Write("d");
            });

            Assert.Equal("a,,,d\r\n", csv);
        }

        [Fact]
        public void BoolIsWrittenLowercase()
        {
            string csv = Write(w =>
            {
                using CsvRowWriter row = w.StartRow();
                row.Write(true);
                row.Write(false);
                row.Write((bool?)null);
            });

            Assert.Equal("true,false,\r\n", csv);
        }

        [Fact]
        public void NumericOverloadsRoundTrip()
        {
            string csv = Write(w =>
            {
                using CsvRowWriter row = w.StartRow();
                row.Write(123);
                row.Write(1234567890123L);
                row.Write(2.75d);
                row.Write(12.5m);
                row.Write((int?)null);
                row.Write((decimal?)8.25m);
            });

            Assert.Equal("123,1234567890123,2.75,12.5,,8.25\r\n", csv);
        }

        [Fact]
        public void NumericFieldContainingDelimiterIsQuoted()
        {
            // Delimiter '.' collides with the decimal point, so the UTF-8 numeric path must quote it.
            string csv = Write(w =>
            {
                using CsvRowWriter row = w.StartRow();
                row.Write("a");
                row.Write(3.5);
                row.Write("b");
            }, options: new CsvWriterOptions { Delimiter = (byte)'.' });

            Assert.Equal("a.\"3.5\".b\r\n", csv);
        }

        [Fact]
        public void SemicolonDelimiterIsRespected()
        {
            string csv = Write(w =>
            {
                using CsvRowWriter row = w.StartRow();
                row.Write("a");
                row.Write("b,c"); // comma no longer special once delimiter is ';'
            }, options: new CsvWriterOptions { Delimiter = (byte)';' });

            Assert.Equal("a;b,c\r\n", csv);
        }

        [Fact]
        public void DisposingWriterEndsAnUnfinishedRow()
        {
            var ms = new MemoryStream();
            using (CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true))
            {
                CsvRowWriter row = writer.StartRow();
                row.Write("a");
                row.Write("b");
                // row intentionally left un-disposed
            }

            Assert.Equal("a,b\r\n", Encoding.UTF8.GetString(ms.ToArray()));
        }

        [Fact]
        public void WrittenCsvRoundTripsThroughCsvReader()
        {
            var date = new DateTime(2026, 7, 2, 0, 0, 0, DateTimeKind.Unspecified);
            var ms = new MemoryStream();
            using (CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true))
            {
                using (CsvRowWriter row = writer.StartRow())
                {
                    row.Write("Name");
                    row.Write("Active");
                    row.Write("Created");
                }
                using (CsvRowWriter row = writer.StartRow())
                {
                    row.Write("has, comma");
                    row.Write(true);
                    row.Write(date);
                }
            }
            ms.Position = 0;

            using IExcelRowReader reader = Excel.FromCsv(ms);
            using IExcelRowEnumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            var header = e.Current;
            Assert.Equal("Name", header[0].GetString());

            Assert.True(e.MoveNext());
            var row2 = e.Current;
            Assert.Equal("has, comma", row2[0].GetString());
            Assert.Equal(CellType.ExcelString, row2[1].Type);
            Assert.Equal("true", row2[1].GetString());
            Assert.False(row2[2].TryGetDateTime(out _)); // CSV cells have no serial date form
            Assert.Equal(date.ToString("O"), row2[2].GetString());
        }

        [Fact]
        public void DelimiterEqualToQuoteThrows()
        {
            var ms = new MemoryStream();
            var options = new CsvWriterOptions { Delimiter = (byte)'"' };

            Assert.Throws<ArgumentException>(() => CsvWriter.Create(ms, options: options));
        }

        [Fact]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP017:Prefer using",
            Justification = "Explicit Dispose() call is the point of this test — asserts the stream closes as a side effect.")]
        public void DisposeClosesStreamWhenNotLeaveOpen()
        {
            var ms = new MemoryStream();
            var writer = CsvWriter.Create(ms, leaveOpen: false);

            writer.Dispose();

            Assert.Throws<ObjectDisposedException>(() => ms.ReadByte());
        }

        [Fact]
        public async Task DisposeAsyncLeavesStreamOpenWhenLeaveOpenTrue()
        {
            using var ms = new MemoryStream();
            var writer = CsvWriter.Create(ms, leaveOpen: true);
            using (CsvRowWriter row = writer.StartRow())
            {
                row.Write("a");
            }

            await writer.DisposeAsync();

            ms.Position = 0;
            Assert.Equal("a\r\n", Encoding.UTF8.GetString(ms.ToArray()));
        }
    }
}
