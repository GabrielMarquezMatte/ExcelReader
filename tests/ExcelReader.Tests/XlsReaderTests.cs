using ExcelReader.Core.Enums;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using System.Collections;

namespace ExcelReader.Tests
{
    public class XlsReaderTests
    {
        private sealed class PersonRow
        {
            [ExcelColumn("Preferred Name")]
            [ExcelColumn("Name")]
            public string? Name { get; set; }
            public int Age { get; set; }
            public DateTime BirthDate { get; set; }
            public bool Active { get; set; }
        }

        [Fact]
        public void ReadsBasicBiff8Cells()
        {
            using var ms = XlsWorkbookBuilder.Build(
                sheets: [("S1", [["Name", "Age", "Active", "Err", "Formula"], ["João", 42, true, new XlsError(0x07), new XlsFormula(12.5)]])]);
            using var reader = Excel.FromXls(ms);

            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            RowAssert(e.Current, ["Name", "Age", "Active", "Err", "Formula"]);
            Assert.True(e.MoveNext());
            var row = e.Current;
            Assert.Equal("João", row[0].GetString());
            Assert.Equal(CellType.Number, row[1].Type);
            Assert.True(row[1].TryParse(null, out int age));
            Assert.Equal(42, age);
            Assert.Equal(CellType.Boolean, row[2].Type);
            Assert.Equal("1", row[2].GetString());
            Assert.Equal(CellType.Error, row[3].Type);
            Assert.Equal("#DIV/0!", row[3].GetString());
            Assert.Equal(CellType.Formula, row[4].Type);
            Assert.Equal("12.5", row[4].GetString());
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void ReadsAdditionalBiff8CellEncodings()
        {
            string longText = new('x', 5000);
            using var ms = XlsWorkbookBuilder.Build(
                sheets:
                [
                    ("S1",
                    [
                        [new XlsUnicodeString("Unicode Ω"), new XlsCompressedBytes([0x80]), new XlsRkInt(123), new XlsMulRk(4, 5), null, new XlsFormulaBool(true), new XlsFormulaError(0x2A), new XlsBlank(), new XlsMulBlank(2), null, longText]
                    ])
                ]);
            using var reader = Excel.FromXls(ms);

            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            var row = e.Current;
            Assert.Equal("Unicode Ω", row[0].GetString());
            Assert.Equal("€", row[1].GetString());
            Assert.Equal("123", row[2].GetString());
            Assert.Equal("4", row[3].GetString());
            Assert.Equal("5", row[4].GetString());
            Assert.Equal("1", row[5].GetString());
            Assert.Equal("#N/A", row[6].GetString());
            Assert.Equal(longText, row[10].GetString());
        }

        [Fact]
        public void ReadsWideRowsRkDoublesAndAllErrorNames()
        {
            object?[] wide = new object?[45];
            for (int i = 0; i < wide.Length; i++)
            {
                wide[i] = i;
            }
            wide[5] = new XlsRkRaw(0x3FF00000);
            wide[10] = new XlsError(0x00);
            wide[11] = new XlsError(0x0F);
            wide[12] = new XlsError(0x17);
            wide[13] = new XlsError(0x1D);
            wide[14] = new XlsError(0x24);
            wide[15] = new XlsError(0x55);

            using var ms = XlsWorkbookBuilder.Build(sheets: [("S1", [wide])]);
            using var reader = Excel.FromXls(ms);

            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            var row = e.Current;
            Assert.Equal(45, row.ColumnCount);
            Assert.Equal("1", row[5].GetString());
            Assert.Equal("#NULL!", row[10].GetString());
            Assert.Equal("#VALUE!", row[11].GetString());
            Assert.Equal("#REF!", row[12].GetString());
            Assert.Equal("#NAME?", row[13].GetString());
            Assert.Equal("#NUM!", row[14].GetString());
            Assert.Equal("#ERR", row[15].GetString());
        }

        [Fact]
        public void OutOfOrderCellsAreSortedBeforeRowIsExposed()
        {
            using var ms = XlsWorkbookBuilder.Build(sheets: [("S1", [[new XlsAt(2, "C"), new XlsAt(0, "A"), new XlsAt(1, "B")]])]);
            using var reader = Excel.FromXls(ms);

            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("A", e.Current[0].GetString());
            Assert.Equal("B", e.Current[1].GetString());
            Assert.Equal("C", e.Current[2].GetString());
        }

        [Fact]
        public void ReadsSharedStringsAndMultipleSheets()
        {
            using var ms = XlsWorkbookBuilder.Build(
                sheets:
                [
                    ("First", [[new XlsSharedString("Ignored")]]),
                    ("Second", [[new XlsSharedString("Header")], [new XlsSharedString("Café")]])
                ]);
            using var reader = Excel.FromXls(ms);

            Assert.Equal(2, reader.SheetCount);
            Assert.True(reader.TryMoveToSheet("second"));
            Assert.Equal("Second", reader.SheetName);
            reader.MoveToSheet(0);
            Assert.Equal("First", reader.SheetName);
            reader.MoveToSheet(1);
            Assert.Equal("Second", reader.SheetName);

            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("Header", e.Current[0].GetString());
            Assert.True(e.MoveNext());
            Assert.Equal("Café", e.Current[0].GetString());
        }

        [Fact]
        public void LargeSharedStringGrowsSharedBuffer()
        {
            string text = new('s', 600);
            using var ms = XlsWorkbookBuilder.Build(sheets: [("S1", [[new XlsSharedString(text)]])]);
            using var reader = Excel.FromXls(ms);

            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(text, e.Current[0].GetString());
        }

        [Fact]
        public void UnicodeSheetNamesAreDecoded()
        {
            using var ms = XlsWorkbookBuilder.Build(sheets: [("Ωmega", [["A"]])]);
            using var reader = Excel.FromXls(ms);

            Assert.Equal("Ωmega", reader.SheetName);
            Assert.True(reader.TryMoveToSheet("ωmega"));
        }

        [Fact]
        public void Date1904IsAppliedByParser()
        {
            var date = new DateTime(1904, 1, 2, 0, 0, 0, DateTimeKind.Unspecified);
            using var ms = XlsWorkbookBuilder.Build(
                date1904: true,
                sheets: [("S1", [["BirthDate"], [new XlsDate(date)]])]);
            using var reader = Excel.FromXls(ms);

            Assert.True(reader.IsDate1904);
            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal(date, result[0].BirthDate);
        }

        [Fact]
        public void CustomDateFormatMarksDateCells()
        {
            var date = new DateTime(2024, 5, 6, 0, 0, 0, DateTimeKind.Unspecified);
            using var ms = XlsWorkbookBuilder.Build(
                customDateFormat: "yyyy-mm-dd",
                sheets: [("S1", [["BirthDate"], [new XlsDate(date)]])]);
            using var reader = Excel.FromXls(ms);

            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Date, e.Current[0].Type);
            Assert.True(e.Current[0].TryGetDateTime(out DateTime parsed));
            Assert.Equal(date, parsed);
        }

        [Fact]
        public void CustomNonDateFormatKeepsNumberCells()
        {
            using var ms = XlsWorkbookBuilder.Build(
                customDateFormat: "0.00",
                sheets: [("S1", [["BirthDate"], [new XlsDate(new DateTime(2024, 5, 6, 0, 0, 0, DateTimeKind.Unspecified))]])]);
            using var reader = Excel.FromXls(ms);

            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Number, e.Current[0].Type);
        }

        [Theory]
        [InlineData("\"yyyy\"")]
        [InlineData("[red]0")]
        [InlineData("\\d0")]
        [InlineData("\"unterminated y")]
        [InlineData("[unterminated y")]
        public void CustomFormatWithoutDateTokensKeepsNumberCells(string format)
        {
            using var ms = XlsWorkbookBuilder.Build(
                customDateFormat: format,
                sheets: [("S1", [["BirthDate"], [new XlsDate(new DateTime(2024, 5, 6, 0, 0, 0, DateTimeKind.Unspecified))]])]);
            using var reader = Excel.FromXls(ms);

            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Number, e.Current[0].Type);
        }

        [Fact]
        public void UnicodeCustomDateFormatMarksDateCells()
        {
            using var ms = XlsWorkbookBuilder.Build(
                customDateFormat: "yyyy Ω",
                sheets: [("S1", [["BirthDate"], [new XlsDate(new DateTime(2024, 5, 6, 0, 0, 0, DateTimeKind.Unspecified))]])]);
            using var reader = Excel.FromXls(ms);

            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Date, e.Current[0].Type);
        }

        [Fact]
        public void TypedParserUsesXlsReaderAndAttributeAliases()
        {
            using var ms = XlsWorkbookBuilder.Build(
                sheets: [("S1", [["Name", "Age", "Active"], ["Ana", 31, false]])]);
            using var reader = Excel.FromXls(ms);

            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Ana", result[0].Name);
            Assert.Equal(31, result[0].Age);
            Assert.False(result[0].Active);
        }

        [Fact]
        public void XlsParserEnumerableSupportsExplicitInterfacesAndResetThrows()
        {
            using var ms = XlsWorkbookBuilder.Build(sheets: [("S1", [["Name"], ["Lua"]])]);
            using var reader = Excel.FromXls(ms);
            var enumerable = new ExcelParser<PersonRow>().Parse(reader);

            using IEnumerator<PersonRow> generic = ((IEnumerable<PersonRow>)enumerable).GetEnumerator();
            Assert.True(generic.MoveNext());
            Assert.Equal("Lua", generic.Current.Name);
            Assert.Throws<NotSupportedException>(generic.Reset);

            IEnumerator nongeneric = ((IEnumerable)enumerable).GetEnumerator();
            try
            {
                Assert.True(nongeneric.MoveNext());
                Assert.IsType<PersonRow>(nongeneric.Current);
            }
            finally
            {
                (nongeneric as IDisposable)?.Dispose();
            }
        }

        [Fact]
        public void XlsParserHeaderRowSkipsLeadingRows()
        {
            using var ms = XlsWorkbookBuilder.Build(sheets: [("S1", [["skip"], ["Name"], ["Sol"]])]);
            using var reader = Excel.FromXls(ms);
            var parser = new ExcelParser<PersonRow>(new ExcelParserConfig { HeaderRow = 2 });

            var result = parser.Parse(reader).ToList();
            Assert.Single(result);
            Assert.Equal("Sol", result[0].Name);
        }

        [Fact]
        public void XlsParserReturnsDefaultRowsWhenHeaderMapsNoColumns()
        {
            using var ms = XlsWorkbookBuilder.Build(sheets: [("S1", [["Unknown"], ["value"]])]);
            using var reader = Excel.FromXls(ms);

            var result = new ExcelParser<PersonRow>().Parse(reader).ToList();
            Assert.Single(result);
            Assert.Null(result[0].Name);
        }

        [Fact]
        public async Task XlsAsyncParserReturnsDefaultRowsWhenHeaderMapsNoColumns()
        {
            using var ms = XlsWorkbookBuilder.Build(sheets: [("S1", [["Unknown"], ["value"]])]);
            using var reader = Excel.FromXls(ms);
            var rows = new List<PersonRow>();

            await foreach (PersonRow row in new ExcelParser<PersonRow>().ParseAsync(reader, TestContext.Current.CancellationToken))
            {
                rows.Add(row);
            }

            Assert.Single(rows);
            Assert.Null(rows[0].Name);
        }

        [Fact]
        public async Task AsyncReaderAndParserMatchSync()
        {
            await using var ms = XlsWorkbookBuilder.Build(
                sheets: [("S1", [["Preferred Name", "Age"], ["Bia", 27]])]);
            await using var reader = await Excel.FromXlsAsync(ms, ct: TestContext.Current.CancellationToken);

            List<PersonRow> rows = [];
            await foreach (PersonRow row in new ExcelParser<PersonRow>().ParseAsync(reader, TestContext.Current.CancellationToken))
            {
                rows.Add(row);
            }

            Assert.Single(rows);
            Assert.Equal("Bia", rows[0].Name);
            Assert.Equal(27, rows[0].Age);
        }

        [Fact]
        public async Task AsyncParserDisposeBeforeMoveNextIsNoop()
        {
            using var ms = XlsWorkbookBuilder.Build(sheets: [("S1", [["Name"], ["Noop"]])]);
            using var reader = Excel.FromXls(ms);

            IAsyncEnumerator<PersonRow> e = new ExcelParser<PersonRow>()
                .ParseAsync(reader, TestContext.Current.CancellationToken)
                .GetAsyncEnumerator(TestContext.Current.CancellationToken);

            Exception? ex = await Record.ExceptionAsync(async () =>
                await e.DisposeAsync());

            Assert.Null(ex);
        }

        [Fact]
        public void CancelledAsyncEnumeratorThrows()
        {
            using var ms = XlsWorkbookBuilder.Build(sheets: [("S1", [["A"]])]);
            using var reader = Excel.FromXls(ms);
            using var cts = new CancellationTokenSource();
            cts.Cancel();

            Assert.Throws<OperationCanceledException>(() => reader.GetAsyncEnumerator(cts.Token));
        }

        [Fact]
        public void SharedStringIndexOutsideTableYieldsEmptyString()
        {
            using var ms = XlsWorkbookBuilder.Build(sheets: [("S1", [[new XlsSharedIndex(99)]])]);
            using var reader = Excel.FromXls(ms);

            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(string.Empty, e.Current[0].GetString());
        }

        [Fact]
        public async Task FileFactoryMethodsOpenXlsFiles()
        {
            string path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xls");
            try
            {
                await File.WriteAllBytesAsync(path, XlsWorkbookBuilder.Build(sheets: [("S1", [["A"], ["B"]])]).ToArray(), TestContext.Current.CancellationToken);

                using (var reader = Excel.FromXlsFile(path))
                {
                    using var e = reader.GetEnumerator();
                    Assert.True(e.MoveNext());
                    Assert.Equal("A", e.Current[0].GetString());
                }

                await using (var reader = await Excel.FromXlsFileAsync(path, TestContext.Current.CancellationToken))
                {
                    await using var e = reader.GetAsyncEnumerator(TestContext.Current.CancellationToken);
                    Assert.True(await e.MoveNextAsync());
                    Assert.Equal("A", e.Current[0].GetString());
                }
            }
            finally
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
        }

        [Fact]
        public void SheetNavigationValidatesIndexAndMissingName()
        {
            using var ms = XlsWorkbookBuilder.Build(sheets: [("S1", [["A"]])]);
            using var reader = Excel.FromXls(ms);

            Assert.False(reader.TryMoveToSheet("Missing"));
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.MoveToSheet(-1));
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.MoveToSheet(1));
        }

        [Fact]
        public void LeaveOpenFalseDisposesSourceStream()
        {
            var stream = new TrackingStream(XlsWorkbookBuilder.Build(sheets: [("S1", [["A"]])]).ToArray());
            using (Excel.FromXls(stream, leaveOpen: false))
            {
                // nothing
            }
            Assert.True(stream.WasDisposed);
        }

        [Fact]
        public void InvalidOleDocumentThrows()
        {
            using MemoryStream ms = new([1, 2, 3, 4]);
            Assert.Throws<InvalidDataException>(() => Excel.FromXls(ms));
        }

        [Fact]
        public void NonBiff8WorkbookThrows()
        {
            using var ms = XlsWorkbookBuilder.BuildNonBiff8();
            Assert.Throws<NotSupportedException>(() => Excel.FromXls(ms));
        }

        [Fact]
        public void EncryptedWorkbookThrows()
        {
            using var ms = XlsWorkbookBuilder.BuildEncrypted();
            Assert.Throws<NotSupportedException>(() => Excel.FromXls(ms));
        }

        [Fact]
        public void WorkbookWithoutSheetsThrows()
        {
            using var ms = XlsWorkbookBuilder.BuildNoSheets();
            Assert.Throws<InvalidDataException>(() => Excel.FromXls(ms));
        }

        [Fact]
        public void BadWorksheetBofThrowsWhenEnumerated()
        {
            using var ms = XlsWorkbookBuilder.BuildBadSheetBof();
            using var reader = Excel.FromXls(ms);
            using var e = reader.GetEnumerator();

            Assert.Throws<NotSupportedException>(() => e.MoveNext());
        }

        [Fact]
        public void ShortAndUnknownSheetRecordsAreSkippedSafely()
        {
            using var ms = XlsWorkbookBuilder.BuildRawSheet(
                includeEof: false,
                (0x00FF, [1, 2]),
                (0x0204, []),
                (0x0204, XlsWorkbookBuilder.RawRowOnly(0)),
                (0x00BD, XlsWorkbookBuilder.RawRowOnly(0)),
                (0x0205, XlsWorkbookBuilder.RawRowOnly(0)),
                (0x0006, XlsWorkbookBuilder.RawRowOnly(0)),
                (0x0204, XlsWorkbookBuilder.RawLabel(0, 0, "A")));
            using var reader = Excel.FromXls(ms);
            using var e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("A", e.Current[0].GetString());
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void ReadsManyRowsSpanningMultipleSectorsWithExactCount()
        {
            const int dataRows = 700;
            object?[][] rows = new object?[dataRows + 1][];
            rows[0] = ["Name", "Age", "Score"];
            for (int r = 0; r < dataRows; r++)
            {
                rows[r + 1] = [$"row{r}", r, r * 1.5];
            }

            using var ms = XlsWorkbookBuilder.Build(sheets: [("S1", rows)]);
            Assert.True(ms.Length > SectorSize * 4, "workbook should span several OLE sectors");
            using var reader = Excel.FromXls(ms);

            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            RowAssert(e.Current, ["Name", "Age", "Score"]);

            int read = 0;
            while (e.MoveNext())
            {
                var row = e.Current;
                Assert.Equal(3, row.ColumnCount);
                Assert.Equal($"row{read}", row[0].GetString());
                Assert.True(row[1].TryParse(null, out int age));
                Assert.Equal(read, age);
                Assert.True(row[2].TryParse(System.Globalization.CultureInfo.InvariantCulture, out double score));
                Assert.Equal(read * 1.5, score);
                read++;
            }
            Assert.Equal(dataRows, read);
        }

        [Fact]
        public void NumericCellsExposeRawDoubleMatchingFormattedText()
        {
            using var ms = XlsWorkbookBuilder.Build(
                sheets: [("S1", [[12.5, 42, new XlsRkInt(123), -7.25]])]);
            using var reader = Excel.FromXls(ms);

            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            var row = e.Current;

            ReadOnlySpan<double> expected = [12.5, 42, 123, -7.25];
            for (int c = 0; c < expected.Length; c++)
            {
                // Raw fast path returns the exact stored double...
                Assert.True(row[c].TryGetDouble(out double raw));
                Assert.Equal(expected[c], raw);
                Assert.True(row[c].TryParse(System.Globalization.CultureInfo.InvariantCulture, out double parsed));
                Assert.Equal(expected[c], parsed);
                // ...and the text representation stays consistent with parsing it back.
                Assert.True(double.TryParse(row[c].GetString(), System.Globalization.CultureInfo.InvariantCulture, out double fromText));
                Assert.Equal(expected[c], fromText);
            }
        }

        [Fact]
        public void SharedStringSplitAcrossContinueBoundaryDecodesCorrectly()
        {
            // SST record ends mid-way through string 1's character array; the CONTINUE record resumes
            // with a fresh grbit byte. The decoder must consume that byte, not read it as a character.
            // string 0 = "AB" (cch=2, compressed); string 1 = "CDEF" split after "CD".
            byte[] firstRegion =
            [
                0x02, 0x00, 0x00, (byte)'A', (byte)'B',       // "AB"
                0x04, 0x00, 0x00, (byte)'C', (byte)'D',       // "CDEF" header + first two chars
            ];
            byte[] continueRegion = [0x00, (byte)'E', (byte)'F']; // grbit + remaining two chars
            byte[] framed = XlsWorkbookBuilder.FrameSstWithContinue(2, 2, firstRegion, continueRegion);

            using var ms = XlsWorkbookBuilder.BuildRawSst(framed, labelSstCount: 2);
            using var reader = Excel.FromXls(ms);

            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("AB", e.Current[0].GetString());
            Assert.Equal("CDEF", e.Current[1].GetString());   // was corrupted before the CONTINUE fix
        }

        [Fact]
        public void AstralCharInUnicodeStringRoundTripsAsValidUtf8()
        {
            // A surrogate pair must encode as one 4-byte UTF-8 scalar, not two 3-byte CESU-8 sequences.
            const string emoji = "A\U0001F600B"; // A, grinning face, B
            using var ms = XlsWorkbookBuilder.Build(
                sheets: [("S1", [[new XlsUnicodeString(emoji)]])]);
            using var reader = Excel.FromXls(ms);

            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(emoji, e.Current[0].GetString());
        }

        [Fact]
        public void CraftedFatSectorCountThrowsInsteadOfAllocating()
        {
            // A bogus FAT sector count in the OLE header must be rejected, not turned into new int[huge].
            using var ms = XlsWorkbookBuilder.BuildPatched(
                XlsWorkbookBuilder.FatSectorCountOffset, XlsWorkbookBuilder.LE32(0x40000000));
            Assert.Throws<InvalidDataException>(() => Excel.FromXls(ms));
        }

        [Fact]
        public async Task WritesAndReadsLongLabelSplitsAcrossContinueRecords()
        {
            // Create a string of 10,000 characters (exceeds 8,224 bytes MaxPayload)
            // with some non-ASCII CP1252 characters to verify compression handles correctly.
            string longCompressed = new string('a', 5000) + "Café" + new string('b', 5000);
            string longWide = new string('Ω', 6000);

            var ms = new MemoryStream();
            await using (XlsWorkbookWriter wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                wb.Start();
                XlsSheetWriter sheet = wb.AddSheet("S1");
                sheet.Start();
                using (XlsRowWriter row = sheet.StartRow())
                {
                    row.Write(longCompressed);
                    row.Write(longWide);
                }
                sheet.End();
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }

            ms.Position = 0;
            using var reader = Excel.FromXls(ms);
            using var e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal(longCompressed, e.Current[0].GetString());
            Assert.Equal(longWide, e.Current[1].GetString());
        }

        private const int SectorSize = 512;

        private static void RowAssert(Core.ValueObjects.Row row, string[] values)
        {
            Assert.Equal(values.Length, row.ColumnCount);
            for (int i = 0; i < values.Length; i++)
            {
                Assert.Equal(values[i], row[i].GetString());
            }
        }

        private sealed class TrackingStream : MemoryStream
        {
            internal TrackingStream(byte[] bytes)
                : base(bytes)
            {
            }

            internal bool WasDisposed { get; private set; }

            protected override void Dispose(bool disposing)
            {
                WasDisposed = true;
                base.Dispose(disposing);
            }
        }
    }
}
