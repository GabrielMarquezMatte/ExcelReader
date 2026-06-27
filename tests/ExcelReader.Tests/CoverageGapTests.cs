using System.Globalization;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Core.Writer;
using B = ExcelReader.Tests.Biff12Build;

namespace ExcelReader.Tests
{
    // Targets coverage gaps left by the format-specific suites: Cell text/float paths,
    // header normalization flags, the typed write overloads of the XLS/XLSB row writers,
    // workbook-writer lifecycle errors, and XLSB reader navigation/disposal.
    public class CoverageGapTests
    {
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        // ISpanFormattable but NOT IConvertible — forces the XLSB writer's format-then-parse fallback.
        private readonly struct Formattable(double value) : ISpanFormattable
        {
            public string ToString(string? format, IFormatProvider? formatProvider)
            {
                return value.ToString(formatProvider);
            }

            public bool TryFormat(Span<char> destination, out int charsWritten, ReadOnlySpan<char> format, IFormatProvider? provider)
            {
                return value.TryFormat(destination, out charsWritten, format, provider);
            }
        }

        // --- Cell ---

        [Fact]
        public void CellTryFormatWritesBinaryNumber()
        {
            var cell = new Cell(CellType.Number, default, 42.5, hasNumber: true, 0);
            Span<byte> buf = stackalloc byte[32];
            Assert.True(cell.TryFormat(buf, out int written));
            Assert.Equal("42.5", System.Text.Encoding.UTF8.GetString(buf[..written]));
        }

        [Fact]
        public void CellTryFormatCopiesTextValue()
        {
            var cell = new Cell(CellType.ExcelString, "hi"u8);
            Span<byte> buf = stackalloc byte[8];
            Assert.True(cell.TryFormat(buf, out int written));
            Assert.Equal("hi", System.Text.Encoding.UTF8.GetString(buf[..written]));
        }

        [Fact]
        public void CellTryFormatReturnsFalseWhenDestinationTooSmall()
        {
            Span<byte> tiny = stackalloc byte[2];
            Assert.False(new Cell(CellType.Number, default, 12345.0, hasNumber: true, 0).TryFormat(tiny, out _));
            Assert.False(new Cell(CellType.ExcelString, "hello"u8).TryFormat(tiny, out _));
        }

        [Fact]
        public void CellTryParseFloatFromBinaryNumber()
        {
            var cell = new Cell(CellType.Number, default, 2.5, hasNumber: true, 0);
            Assert.True(cell.TryParse(Inv, out float f));
            Assert.Equal(2.5f, f);
        }

        // --- HeaderNormalization ---

        [Fact]
        public void HeaderNormalizationNoneReturnsValueUnchanged()
        {
            Assert.Equal("  Café  ", HeaderNormalization.None.Apply("  Café  "));
        }

        [Fact]
        public void HeaderNormalizationCollapsesWhitespace()
        {
            const HeaderNormalization norm = HeaderNormalization.Trim | HeaderNormalization.CollapseSpaces;
            Assert.Equal("a b c", norm.Apply("  a   b\t\r\nc  "));
        }

        [Fact]
        public void HeaderNormalizationRemovesDiacritics()
        {
            Assert.Equal("Cafe ond", HeaderNormalization.RemoveDiacritics.Apply("Café ônd"));
        }

        // --- XLSB row writer typed overloads ---

        private static async Task<MemoryStream> WriteXlsbAsync(Func<XlsbWorkbookWriter, Task> build, bool date1904 = false)
        {
            var ms = new MemoryStream();
            await using var wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, date1904: date1904, ct: TestContext.Current.CancellationToken);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            await build(wb);
            await wb.EndAsync(TestContext.Current.CancellationToken);
            ms.Position = 0;
            return ms;
        }

        [Fact]
        public async Task XlsbRowWriterWritesNullableAndGenericValues()
        {
            var date = new DateTime(2001, 2, 3, 0, 0, 0, DateTimeKind.Unspecified);
            await using MemoryStream ms = await WriteXlsbAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken);
                row.Write((bool?)true);          // col 0
                row.Write((bool?)null);           // col 1 → empty
                row.Write((DateTime?)date);       // col 2
                row.Write((DateTime?)null);       // col 3 → empty
                row.Write((int?)7);               // col 4
                row.Write((int?)null);            // col 5 → empty
                row.Write((double?)2.5);          // col 6
                row.Write(new Formattable(3.5));  // col 7 → non-IConvertible fallback
            });

            await using var reader = Excel.FromXlsb(ms);
            using XlsbReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Boolean, e.Current[0].Type);
            Assert.Equal("1", e.Current[0].GetString());
            Assert.Equal(CellType.Empty, e.Current[1].Type);
            Assert.Equal(CellType.Date, e.Current[2].Type);
            Assert.True(e.Current[2].TryGetDateTime(reader.IsDate1904, out DateTime parsed));
            Assert.Equal(date, parsed);
            Assert.Equal(CellType.Empty, e.Current[3].Type);
            Assert.True(e.Current[4].TryGetDouble(out double seven));
            Assert.Equal(7.0, seven);
            Assert.Equal(CellType.Empty, e.Current[5].Type);
            Assert.True(e.Current[6].TryGetDouble(out double half));
            Assert.Equal(2.5, half);
            Assert.True(e.Current[7].TryGetDouble(out double frac));
            Assert.Equal(3.5, frac);
        }

        [Fact]
        public async Task XlsbRowWriterSkipNegativeThrows()
        {
            await using MemoryStream ms = await WriteXlsbAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken);
                Assert.Throws<ArgumentOutOfRangeException>(() => row.Skip(-1));
                row.Write("x");
            });

            await using var reader = Excel.FromXlsb(ms);
            using XlsbReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("x", e.Current[0].GetString());
        }

        // --- XLS row writer nullable overloads ---

        [Fact]
        public async Task XlsRowWriterWritesNullableValues()
        {
            var date = new DateTime(1999, 12, 31, 0, 0, 0, DateTimeKind.Unspecified);
            var ms = new MemoryStream();
            await using (XlsWorkbookWriter wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                wb.Start();
                XlsSheetWriter sheet = wb.AddSheet("S1");
                sheet.Start();
                using (XlsRowWriter row = sheet.StartRow())
                {
                    row.Write((bool?)true);     // col 0
                    row.Write((bool?)null);     // col 1 → empty
                    row.Write((DateTime?)date); // col 2
                    row.Write((DateTime?)null); // col 3 → empty
                    row.Write((int?)11);        // col 4
                    row.Write((int?)null);      // col 5 → empty
                }
                sheet.End();
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }

            ms.Position = 0;
            using var reader = Excel.FromXls(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Boolean, e.Current[0].Type);
            Assert.Equal(CellType.Empty, e.Current[1].Type);
            Assert.Equal(CellType.Date, e.Current[2].Type);
            Assert.True(e.Current[2].TryGetDateTime(out DateTime parsed));
            Assert.Equal(date, parsed);
            Assert.Equal(CellType.Empty, e.Current[3].Type);
            Assert.True(e.Current[4].TryParse(Inv, out int eleven));
            Assert.Equal(11, eleven);
            Assert.Equal(CellType.Empty, e.Current[5].Type);
        }

        // --- XLSB workbook writer lifecycle ---

        [Fact]
        public async Task XlsbWorkbookWriterLifecycleErrors()
        {
            await using var wb = await XlsbWorkbookWriter.CreateAsync(new MemoryStream(), ct: TestContext.Current.CancellationToken);
            Assert.Throws<InvalidOperationException>(() => wb.AddSheet("S1")); // before Start

            await wb.StartAsync(TestContext.Current.CancellationToken);
            await Assert.ThrowsAsync<InvalidOperationException>(async () => await wb.StartAsync(TestContext.Current.CancellationToken));
            Assert.Throws<ArgumentException>(() => wb.AddSheet(""));
            XlsbSheetWriter s1 = wb.AddSheet("S1");
            Assert.Throws<InvalidOperationException>(() => wb.AddSheet("S2")); // previous not ended

            // End the sheet so the writer disposes cleanly (it needs at least one registered sheet).
            await s1.StartAsync(TestContext.Current.CancellationToken);
            await s1.EndAsync(TestContext.Current.CancellationToken);
        }

        [Fact]
        public async Task XlsbWorkbookWriterEndWithoutSheetsThrows()
        {
            await using var wb = await XlsbWorkbookWriter.CreateAsync(new MemoryStream(), ct: TestContext.Current.CancellationToken);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            await Assert.ThrowsAsync<InvalidOperationException>(async () => await wb.EndAsync(TestContext.Current.CancellationToken));
        }

        [Fact]
        public async Task XlsbWorkbookWriterFlushDoesNotThrow()
        {
            await using var wb = await XlsbWorkbookWriter.CreateAsync(new MemoryStream(), leaveOpen: true, ct: TestContext.Current.CancellationToken);
            await wb.StartAsync(TestContext.Current.CancellationToken);
            await wb.FlushAsync(TestContext.Current.CancellationToken);
        }

        [Fact]
        public async Task XlsbWorkbookWriterDisposeFinalizesActiveSheetAndClosesStream()
        {
            var ms = new TrackingStream();
            await using (var wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: false, ct: TestContext.Current.CancellationToken))
            {
                await wb.StartAsync(TestContext.Current.CancellationToken);
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken);
                row.Write("v");
                // No EndAsync on sheet or workbook — disposal must finalize both.
            }
            Assert.True(ms.Disposed);
        }

        [Fact]
        public async Task XlsbWorkbookWriterDisposeWithoutStartDisposesStream()
        {
            var ms = new TrackingStream();
            await using (await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: false, ct: TestContext.Current.CancellationToken))
            {
                // Never started.
            }
            Assert.True(ms.Disposed);
        }

        [Fact]
        public async Task XlsbSheetWriterDisposeAutoEnds()
        {
            await using MemoryStream ms = await WriteXlsbAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                {
                    row.Write("auto");
                }
                await sheet.DisposeAsync(); // no explicit EndAsync
            });

            await using var reader = Excel.FromXlsb(ms);
            using XlsbReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("auto", e.Current[0].GetString());
        }

        // --- XLS workbook writer lifecycle ---

        [Fact]
        public async Task XlsWorkbookWriterLifecycleErrors()
        {
            await using var wb = XlsWorkbookWriter.Create(new MemoryStream());
            Assert.Throws<InvalidOperationException>(() => wb.AddSheet("S1")); // before Start
            await Assert.ThrowsAsync<InvalidOperationException>(async () => await wb.EndAsync(TestContext.Current.CancellationToken)); // before Start

            wb.Start();
            Assert.Throws<InvalidOperationException>(wb.Start); // double Start
            Assert.Throws<ArgumentException>(() => wb.AddSheet(""));
            wb.AddSheet("S1");
            Assert.Throws<InvalidOperationException>(() => wb.AddSheet("S2")); // previous not ended
        }

        [Fact]
        public async Task XlsWorkbookWriterFlushDoesNotThrow()
        {
            await using var wb = XlsWorkbookWriter.Create(new MemoryStream(), leaveOpen: true);
            wb.Start();
            await wb.FlushAsync(TestContext.Current.CancellationToken);
        }

        [Fact]
        public async Task XlsWorkbookWriterDisposeFinalizesAndClosesStream()
        {
            var ms = new TrackingStream();
            await using (var wb = XlsWorkbookWriter.Create(ms, leaveOpen: false))
            {
                wb.Start();
                XlsSheetWriter sheet = wb.AddSheet("S1");
                sheet.Start();
                using XlsRowWriter row = sheet.StartRow();
                row.Write("v");
                // No End on sheet or workbook — disposal must finalize and close the stream.
            }
            Assert.True(ms.Disposed);
        }

        [Fact]
        public async Task XlsWorkbookWriterDoubleDisposeIsNoOp()
        {
            var wb = XlsWorkbookWriter.Create(new MemoryStream(), leaveOpen: true);
            wb.Start();
            await wb.DisposeAsync();
            await wb.DisposeAsync(); // second dispose returns immediately
        }

        // --- XLSB reader navigation & disposal ---

        [Fact]
        public async Task XlsbReaderNavigatesMultipleSheets()
        {
            await using MemoryStream ms = await WriteXlsbAsync(async wb =>
            {
                XlsbSheetWriter first = wb.AddSheet("First");
                await first.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter r = await first.StartRowAsync(TestContext.Current.CancellationToken)) { r.Write("a"); }
                await first.EndAsync(TestContext.Current.CancellationToken);

                XlsbSheetWriter second = wb.AddSheet("Second");
                await second.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter r = await second.StartRowAsync(TestContext.Current.CancellationToken)) { r.Write("b"); }
                await second.EndAsync(TestContext.Current.CancellationToken);
            });

            await using var reader = Excel.FromXlsb(ms);
            Assert.Equal(2, reader.SheetCount);
            Assert.Equal("First", reader.SheetName);

            reader.MoveToSheet(1);
            Assert.Equal("Second", reader.SheetName);
            Assert.True(reader.TryMoveToSheet("First"));
            Assert.Equal("First", reader.SheetName);
            Assert.False(reader.TryMoveToSheet("Missing"));
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.MoveToSheet(5));
        }

        [Fact]
        public async Task XlsbReaderRowReaderInterfaceEnumeratesSyncAndAsync()
        {
            await using MemoryStream ms = await WriteXlsbAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter r = await sheet.StartRowAsync(TestContext.Current.CancellationToken)) { r.Write("iface"); }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            });

            await using var reader = Excel.FromXlsb(ms);
            IExcelRowReader rowReader = reader;

            using (IExcelRowEnumerator sync = rowReader.GetEnumerator())
            {
                Assert.True(sync.MoveNext());
                Assert.Equal("iface", sync.Current[0].GetString());
            }

            await using var asyncEnumerator = await rowReader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);
            Assert.True(await asyncEnumerator.MoveNextAsync());
            Assert.Equal("iface", asyncEnumerator.Current[0].GetString());
        }

        [Fact]
        public async Task XlsbReaderLeaveOpenFalseClosesStream()
        {
            byte[] bytes;
            await using (MemoryStream src = await WriteXlsbAsync(async wb =>
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(TestContext.Current.CancellationToken);
                await using (XlsbRowWriter r = await sheet.StartRowAsync(TestContext.Current.CancellationToken)) { r.Write("x"); }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
            }))
            {
                bytes = src.ToArray();
            }

            var tracking = new TrackingStream();
            tracking.Write(bytes);
            tracking.Position = 0;
            using (Excel.FromXlsb(tracking, leaveOpen: false))
            {
                // reader takes ownership
            }
            Assert.True(tracking.Disposed);
        }

        [Fact]
        public void XlsbReaderOpenFailureDisposesStream()
        {
            var tracking = new TrackingStream();
            using (var zip = new System.IO.Compression.ZipArchive(tracking, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            {
                // A ZIP with no workbook part → ParseSheets yields zero sheets → throws.
                zip.CreateEntry("ignored.txt");
            }
            tracking.Position = 0;

            Assert.Throws<InvalidDataException>(() => Excel.FromXlsb(tracking, leaveOpen: false));
            Assert.True(tracking.Disposed);
        }

        [Fact]
        public async Task XlsbReaderOpenFailureAsyncDisposesStream()
        {
            var tracking = new TrackingStream();
            using (var zip = new System.IO.Compression.ZipArchive(tracking, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            {
                zip.CreateEntry("ignored.txt");
            }
            tracking.Position = 0;

            await Assert.ThrowsAsync<InvalidDataException>(async () =>
                await Excel.FromXlsbAsync(tracking, leaveOpen: false, ct: TestContext.Current.CancellationToken));
            Assert.True(tracking.Disposed);
        }

        [Fact]
        public void XlsbReaderSharedAtOutOfRangeYieldsEmpty()
        {
            // CellIsst referencing index 5 when only 1 shared string exists → SharedAt returns (0,0).
            var reader = new XlsbReader(
                sharedFlat: System.Text.Encoding.UTF8.GetBytes("only"),
                sharedOffsets: [0, 4],
                styleIsDate: [],
                date1904: false);
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellIsst, B.CellIsst(0, 0, 5)),
            ];
            using XlsbReader.Enumerator e = reader.GetEnumerator(new MemoryStream(sheet));
            Assert.True(e.MoveNext());
            Assert.Equal(string.Empty, e.Current[0].GetString());
        }

        // Tracks whether the stream was disposed; used to assert leaveOpen semantics.
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
