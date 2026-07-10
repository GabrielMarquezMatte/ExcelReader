using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Parser;
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

        // IUtf8SpanFormattable with a fully controllable formatted length — used to force
        // CsvRowWriter's overflow paths (WriteUtf8Field's stackalloc-too-small fallback and
        // WriteUtf8FieldSlow's rented-buffer-grows-until-it-fits loop), which no BCL numeric
        // type is long enough to reach on its own.
        private readonly struct HugeUtf8Formattable(int length) : IUtf8SpanFormattable
        {
            [SuppressMessage("Major Code Smell", "S1172:Unused method parameters should be removed",
                Justification = "Signature required by IFormattable; this test type ignores format/culture.")]
            public string ToString(string? format, IFormatProvider? formatProvider)
            {
                return new string('x', length);
            }

            public bool TryFormat(Span<byte> utf8Destination, out int bytesWritten, ReadOnlySpan<char> format, IFormatProvider? provider)
            {
                if (utf8Destination.Length < length)
                {
                    bytesWritten = 0;
                    return false;
                }
                utf8Destination[..length].Fill((byte)'x');
                bytesWritten = length;
                return true;
            }
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
        public void StartRowWhileRowActiveThrows()
        {
            var ms = new MemoryStream();
            using CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true);
            CsvRowWriter row = writer.StartRow(); // intentionally left un-disposed
            row.Write("a");

            Assert.Throws<InvalidOperationException>(() => writer.StartRow());
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
        public void NullableBoolWithValueIsWrittenLowercase()
        {
            string csv = Write(w =>
            {
                using CsvRowWriter row = w.StartRow();
                row.Write((bool?)true);
                row.Write((bool?)false);
            });

            Assert.Equal("true,false\r\n", csv);
        }

        [Fact]
        public void NullableDateTimeWithValueIsWrittenIso8601()
        {
            var date = new DateTime(2026, 7, 8, 0, 0, 0, DateTimeKind.Unspecified);
            string csv = Write(w =>
            {
                using CsvRowWriter row = w.StartRow();
                row.Write((DateTime?)date);
                row.Write((DateTime?)null);
            });

            Assert.Equal($"{date:O},\r\n", csv);
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
        public void NumericFieldContainingQuoteByteIsEscaped()
        {
            // Quote='5' makes the digit '5' inside a numeric value collide with the quote byte,
            // forcing WriteFieldBytes's escape-doubling loop (the UTF-8/numeric sibling of
            // WriteStringField's, exercised separately by FieldContainingQuoteIsEscaped).
            var options = new CsvWriterOptions { Quote = (byte)'5' };
            var ms = new MemoryStream();
            using (CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true, options: options))
            {
                using CsvRowWriter row = writer.StartRow();
                row.Write(155);
            }
            ms.Position = 0;

            using IExcelRowReader reader = Excel.FromCsv(ms, options: new CsvReaderOptions { Quote = (byte)'5' });
            using IExcelRowEnumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("155", e.Current[0].GetString());
        }

        [Fact]
        public void PathologicallyLongFormattableValueUsesOverflowPath()
        {
            // 300 bytes: too big for WriteUtf8Field's 64-byte stack buffer, and too big for
            // WriteUtf8FieldSlow's initial 256-byte rental, forcing it to grow once and retry.
            var value = new HugeUtf8Formattable(300);
            string csv = Write(w =>
            {
                using CsvRowWriter row = w.StartRow();
                row.Write(value);
            });

            Assert.Equal(new string('x', 300) + "\r\n", csv);
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
        public void RowDisposedTwiceIsNoOp()
        {
            var ms = new MemoryStream();
            using CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true);
            CsvRowWriter row = writer.StartRow();
            row.Write("a");
            row.Dispose();

            Exception? ex = Record.Exception(row.Dispose);

            Assert.Null(ex);
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

        private sealed record CsvPerson
        {
            public string? Name { get; init; }
            public int Age { get; init; }
            public decimal Balance { get; init; }
            public bool Active { get; init; }
            public DateOnly BirthDate { get; init; }
        }

        [Fact]
        public async Task RecordWriterRoundTripsThroughCsv()
        {
            var people = new[]
            {
                new CsvPerson { Name = "has, comma", Age = 1, Balance = 10.5m, Active = true, BirthDate = new DateOnly(2000, 1, 2) },
                new CsvPerson { Name = "plain", Age = 2, Balance = 20m, Active = false, BirthDate = new DateOnly(2001, 3, 4) },
            };

            var ms = new MemoryStream();
            var ct = TestContext.Current.CancellationToken;
            await using (var writer = RecordWriter.CreateCsv(ms, leaveOpen: true))
            {
                await writer.WriteSheetAsync("People", people, ct);
            }
            ms.Position = 0;

            // Concrete CsvReader (not IExcelRowReader) so Parse picks the CSV text-date overload.
            using var reader = Excel.FromCsv(ms);
            var parsed = new ExcelParser<CsvPerson>().Parse(reader).ToList();

            Assert.Equal(2, parsed.Count);
            Assert.Equal("has, comma", parsed[0].Name);
            Assert.Equal(10.5m, parsed[0].Balance);
            Assert.True(parsed[0].Active);
            Assert.Equal(new DateOnly(2000, 1, 2), parsed[0].BirthDate);
            Assert.Equal(2, parsed[1].Age);
            Assert.False(parsed[1].Active);
        }

        [Fact]
        public async Task RecordWriterRejectsSecondSheet()
        {
            var ms = new MemoryStream();
            var ct = TestContext.Current.CancellationToken;
            await using var writer = RecordWriter.CreateCsv(ms, leaveOpen: true);
            await writer.WriteSheetAsync("One", Array.Empty<CsvPerson>(), ct);

            await Assert.ThrowsAsync<InvalidOperationException>(
                async () => await writer.WriteSheetAsync("Two", Array.Empty<CsvPerson>(), ct));
        }

        [Fact]
        public void DelimiterEqualToQuoteThrows()
        {
            var ms = new MemoryStream();
            var options = new CsvWriterOptions { Delimiter = (byte)'"' };

            Assert.Throws<ArgumentException>(() => CsvWriter.Create(ms, options: options));
        }

        [Theory]
        [InlineData((byte)'\r')]
        [InlineData((byte)'\n')]
        public void DelimiterCarriageReturnOrLineFeedThrows(byte delimiter)
        {
            var ms = new MemoryStream();
            var options = new CsvWriterOptions { Delimiter = delimiter };

            Assert.Throws<ArgumentException>(() => CsvWriter.Create(ms, options: options));
        }

        [Theory]
        [InlineData((byte)'\r')]
        [InlineData((byte)'\n')]
        public void QuoteCarriageReturnOrLineFeedThrows(byte quote)
        {
            var ms = new MemoryStream();
            var options = new CsvWriterOptions { Quote = quote };

            Assert.Throws<ArgumentException>(() => CsvWriter.Create(ms, options: options));
        }

        [Fact]
        public void LargeFieldTriggersAutomaticFlushBeforeDispose()
        {
            var ms = new MemoryStream();
            string bigField = new('a', 2 * 1024 * 1024); // exceeds the 1 MB auto-flush threshold
            using (CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true))
            {
                using (CsvRowWriter row = writer.StartRow())
                {
                    row.Write(bigField);
                }
                // EndRow's threshold check must have flushed to the stream already, before Dispose.
                Assert.True(ms.Length > 0);
            }

            Assert.Equal(bigField + "\r\n", Encoding.UTF8.GetString(ms.ToArray()));
        }

        [Fact]
        public void FlushWritesBufferedDataToStream()
        {
            var ms = new MemoryStream();
            using CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true);
            using (CsvRowWriter row = writer.StartRow())
            {
                row.Write("a");
            }
            Assert.Equal(0, ms.Length); // still buffered, not yet flushed to the stream

            writer.Flush();

            Assert.Equal("a\r\n", Encoding.UTF8.GetString(ms.ToArray()));
        }

        [Fact]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP017:Prefer using",
            Justification = "Explicit Dispose() call is the point of this test — Flush must throw afterward.")]
        public void FlushThrowsAfterDispose()
        {
            var ms = new MemoryStream();
            CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true);
            writer.Dispose();

            Assert.Throws<ObjectDisposedException>(writer.Flush);
        }

        [Fact]
        public async Task FlushAsyncWritesBufferedDataToStream()
        {
            var ms = new MemoryStream();
            await using CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true);
            using (CsvRowWriter row = writer.StartRow())
            {
                row.Write("a");
            }
            Assert.Equal(0, ms.Length);

            await writer.FlushAsync(TestContext.Current.CancellationToken);

            Assert.Equal("a\r\n", Encoding.UTF8.GetString(ms.ToArray()));
        }

        [Fact]
        public async Task FlushAsyncWithNothingBufferedDoesNotThrow()
        {
            var ms = new MemoryStream();
            await using CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true);

            await writer.FlushAsync(TestContext.Current.CancellationToken);

            Assert.Equal(0, ms.Length);
        }

        [Fact]
        public async Task FlushAsyncThrowsAfterDispose()
        {
            var ms = new MemoryStream();
            CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true);
            await writer.DisposeAsync();

            await Assert.ThrowsAsync<ObjectDisposedException>(async () =>
                await writer.FlushAsync(TestContext.Current.CancellationToken));
        }

        [Fact]
        public async Task FlushAsyncThrowsWhenCancellationRequested()
        {
            var ms = new MemoryStream();
            await using CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true);

            await Assert.ThrowsAsync<OperationCanceledException>(async () =>
                await writer.FlushAsync(new CancellationToken(canceled: true)));
        }

        [Fact]
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP017:Prefer using",
            Justification = "Explicit Dispose() calls are the point of this test — a second call must be a no-op.")]
        public void DisposeCalledTwiceIsNoOp()
        {
            var ms = new MemoryStream();
            CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true);
            writer.Dispose();

            Exception? ex = Record.Exception(writer.Dispose);

            Assert.Null(ex);
        }

        [Fact]
        public async Task DisposeAsyncCalledTwiceIsNoOp()
        {
            var ms = new MemoryStream();
            CsvWriter writer = CsvWriter.Create(ms, leaveOpen: true);
            await writer.DisposeAsync();

            Exception? ex = await Record.ExceptionAsync(async () => await writer.DisposeAsync());

            Assert.Null(ex);
        }

        [Fact]
        public async Task DisposeAsyncClosesStreamWhenNotLeaveOpen()
        {
            var ms = new MemoryStream();
            CsvWriter writer = CsvWriter.Create(ms, leaveOpen: false);

            await writer.DisposeAsync();

            Assert.Throws<ObjectDisposedException>(() => ms.ReadByte());
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
