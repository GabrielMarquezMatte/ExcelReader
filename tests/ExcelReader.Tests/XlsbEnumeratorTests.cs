using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using B = ExcelReader.Tests.Biff12Build;

namespace ExcelReader.Tests
{
    public class XlsbEnumeratorTests
    {
        private static XlsbReader BlankReader()
        {
            return new(sharedFlat: [], sharedOffsets: [0], styleIsDate: [], date1904: false);
        }


        private static XlsbReader ReaderWithShared(params string[] strings)
        {
            byte[] flat = Encoding.UTF8.GetBytes(string.Concat(strings));
            int[] offsets = new int[strings.Length + 1];
            int pos = 0;
            for (int i = 0; i < strings.Length; i++)
            {
                offsets[i] = pos;
                pos += Encoding.UTF8.GetByteCount(strings[i]);
            }
            offsets[strings.Length] = pos;
            return new XlsbReader(flat, offsets, styleIsDate: [], date1904: false);
        }

        private static XlsbReader.Enumerator Open(XlsbReader reader, byte[] sheetBin, long entryLength = 0, CancellationToken? ct = null)
        {
            return new(reader, new MemoryStream(sheetBin), entryLength, ct ?? TestContext.Current.CancellationToken);
        }

        // --- Basic enumeration ---


        [Fact]
        public void EmptySheetYieldsNoRows()
        {
            using var e = Open(BlankReader(), []);
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void EmptyRowIsSkipped()
        {
            byte[] sheet = [.. B.Record(Brt.RowHdr), .. B.Record(Brt.EndSheetData)];
            using var e = Open(BlankReader(), sheet);
            Assert.False(e.MoveNext());
        }

        [Fact]
        public void EndSheetDataTerminatesBeforeSubsequentRecords()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.EndSheetData),
                // These must not be processed.
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellBool, B.CellBool(0, 0, true)),
            ];
            using var e = Open(BlankReader(), sheet);
            Assert.False(e.MoveNext());
        }

        // --- Cell type decoding ---

        [Fact]
        public void CellRkDecodesNumber()
        {
            const uint rk = (42u << 2) | 0x02; // integer 42
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellRk, B.CellRk(0, 0, rk)),
            ];
            using var e = Open(BlankReader(), sheet);
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Number, e.Current[0].Type);
            Assert.True(e.Current[0].TryGetDouble(out double val));
            Assert.Equal(42.0, val);
        }

        [Fact]
        public void CellRealDecodesDouble()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellReal, B.CellReal(0, 0, 3.14)),
            ];
            using var e = Open(BlankReader(), sheet);
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Number, e.Current[0].Type);
            Assert.True(e.Current[0].TryGetDouble(out double val));
            Assert.Equal(3.14, val);
        }

        [Fact]
        public void CellIsstResolvesSharedString()
        {
            var reader = ReaderWithShared("Hello", "World");
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellIsst, B.CellIsst(0, 0, 0)), // "Hello"
                .. B.Record(Brt.CellIsst, B.CellIsst(1, 0, 1)), // "World"
            ];
            using var e = Open(reader, sheet);
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.ExcelString, e.Current[0].Type);
            Assert.Equal("Hello", e.Current[0].GetString());
            Assert.Equal("World", e.Current[1].GetString());
        }

        [Fact]
        public void CellStDecodesInlineString()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellSt, B.CellSt(0, 0, "Café")),
            ];
            using var e = Open(BlankReader(), sheet);
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.ExcelString, e.Current[0].Type);
            Assert.Equal("Café", e.Current[0].GetString());
        }

        [Fact]
        public void CellBoolDecodesTrueAndFalse()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellBool, B.CellBool(0, 0, true)),
                .. B.Record(Brt.CellBool, B.CellBool(1, 0, false)),
            ];
            using var e = Open(BlankReader(), sheet);
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Boolean, e.Current[0].Type);
            Assert.Equal("1", e.Current[0].GetString());
            Assert.Equal(CellType.Boolean, e.Current[1].Type);
            Assert.Equal("0", e.Current[1].GetString());
        }

        [Theory]
        [InlineData(0x00, "#NULL!")]
        [InlineData(0x07, "#DIV/0!")]
        [InlineData(0x0F, "#VALUE!")]
        [InlineData(0x17, "#REF!")]
        [InlineData(0x1D, "#NAME?")]
        [InlineData(0x24, "#NUM!")]
        [InlineData(0x2A, "#N/A")]
        public void CellErrorDecodesErrorText(byte code, string expected)
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellError, B.CellError(0, 0, code)),
            ];
            using var e = Open(BlankReader(), sheet);
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Error, e.Current[0].Type);
            Assert.Equal(expected, e.Current[0].GetString());
        }

        [Fact]
        public void CellBlankContributesNoValue()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellBlank, [.. B.U32(0), .. B.U32(0)]),
                .. B.Record(Brt.CellBool, B.CellBool(1, 0, true)),
            ];
            using var e = Open(BlankReader(), sheet);
            Assert.True(e.MoveNext());
            // Column 1 is the last populated; column 0 is blank (no CellDesc for it).
            Assert.Equal(CellType.Empty, e.Current[0].Type);
            Assert.Equal(CellType.Boolean, e.Current[1].Type);
        }

        // --- Date style ---

        [Fact]
        public void DateStyleMapsRealToDateType()
        {
            var reader = new XlsbReader([], [0], styleIsDate: [false, true], date1904: false);
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellReal, B.CellReal(0, 1, 45000.0)), // style=1 → date
                .. B.Record(Brt.CellReal, B.CellReal(1, 0, 45000.0)), // style=0 → number
            ];
            using var e = Open(reader, sheet);
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Date, e.Current[0].Type);
            Assert.Equal(CellType.Number, e.Current[1].Type);
        }

        // --- Multi-row ---

        [Fact]
        public void MultipleRowsEnumeratedInOrder()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellBool, B.CellBool(0, 0, true)),
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellSt, B.CellSt(0, 0, "row2")),
                .. B.Record(Brt.EndSheetData),
            ];
            using var e = Open(BlankReader(), sheet);

            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Boolean, e.Current[0].Type);

            Assert.True(e.MoveNext());
            Assert.Equal("row2", e.Current[0].GetString());

            Assert.False(e.MoveNext());
        }

        [Fact]
        public void EmptyRowsBetweenDataRowsAreSkipped()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellBool, B.CellBool(0, 0, true)),
                .. B.Record(Brt.RowHdr), // empty — must be skipped
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellBool, B.CellBool(0, 0, false)),
            ];
            using var e = Open(BlankReader(), sheet);

            Assert.True(e.MoveNext());
            Assert.Equal("1", e.Current[0].GetString());

            Assert.True(e.MoveNext());
            Assert.Equal("0", e.Current[0].GetString());

            Assert.False(e.MoveNext());
        }

        // --- Async path ---

        [Fact]
        public async Task AsyncMoveNextEnumeratesRows()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellSt, B.CellSt(0, 0, "async")),
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellRk, B.CellRk(0, 0, (7u << 2) | 0x02)), // integer 7
            ];
            await using var e = Open(BlankReader(), sheet);

            Assert.True(await e.MoveNextAsync());
            Assert.Equal("async", e.Current[0].GetString());

            Assert.True(await e.MoveNextAsync());
            Assert.True(e.Current[0].TryGetDouble(out double val));
            Assert.Equal(7.0, val);

            Assert.False(await e.MoveNextAsync());
        }

        // --- Formula cells (cached result) ---

        [Fact]
        public void FmlaNumDecodesCachedNumber()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.FmlaNum, B.CellReal(0, 0, 2.5)),
            ];
            using var e = Open(BlankReader(), sheet);
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Number, e.Current[0].Type);
            Assert.True(e.Current[0].TryGetDouble(out double val));
            Assert.Equal(2.5, val);
        }

        [Fact]
        public void FmlaStringDecodesCachedString()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.FmlaString, B.CellSt(0, 0, "formula")),
            ];
            using var e = Open(BlankReader(), sheet);
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.ExcelString, e.Current[0].Type);
            Assert.Equal("formula", e.Current[0].GetString());
        }

        [Fact]
        public void FmlaBoolDecodesCachedBool()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.FmlaBool, B.CellBool(0, 0, true)),
            ];
            using var e = Open(BlankReader(), sheet);
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Boolean, e.Current[0].Type);
            Assert.Equal("1", e.Current[0].GetString());
        }

        [Fact]
        public void FmlaErrorDecodesCachedError()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.FmlaError, B.CellError(0, 0, 0x2A)),
            ];
            using var e = Open(BlankReader(), sheet);
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Error, e.Current[0].Type);
            Assert.Equal("#N/A", e.Current[0].GetString());
        }

        [Fact]
        public void CellRStringDecodesRichInlineString()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellRString, B.CellRString(0, 0, cRun: 0, "rich")),
            ];
            using var e = Open(BlankReader(), sheet);
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.ExcelString, e.Current[0].Type);
            Assert.Equal("rich", e.Current[0].GetString());
        }

        // --- Malformed cell records ---

        [Fact]
        public void CellRecordShorterThanColStyleHeaderIsIgnored()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellBool, B.U32(0)), // only col, no style/value: length < 8
                .. B.Record(Brt.CellBool, B.CellBool(1, 0, true)),
            ];
            using var e = Open(BlankReader(), sheet);
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Empty, e.Current[0].Type);
            Assert.Equal(CellType.Boolean, e.Current[1].Type);
        }

        [Fact]
        public void CellRecordTooShortForItsValueIsIgnored()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                // col + style only, no bool value byte: length == 8, guard on CellBool needs >= 9.
                .. B.Record(Brt.CellBool, [.. B.U32(0), .. B.U32(0)]),
                .. B.Record(Brt.CellBool, B.CellBool(1, 0, false)),
            ];
            using var e = Open(BlankReader(), sheet);
            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Empty, e.Current[0].Type);
            Assert.Equal(CellType.Boolean, e.Current[1].Type);
        }

        // --- Truncated stream ---

        [Fact]
        public void TruncatedRecordBeforeAnyRowHdrThrows()
        {
            byte[] full = [.. B.Record(Brt.CellBool, B.CellBool(0, 0, true))];
            byte[] truncated = full[..^1];
            using var e = Open(BlankReader(), truncated);
            Assert.Throws<InvalidDataException>(() => e.MoveNext());
        }

        [Fact]
        public void TruncatedCellRecordAfterRowHdrThrows()
        {
            byte[] full = [.. B.Record(Brt.RowHdr), .. B.Record(Brt.CellBool, B.CellBool(0, 0, true))];
            byte[] truncated = full[..^1];
            using var e = Open(BlankReader(), truncated);
            Assert.Throws<InvalidDataException>(() => e.MoveNext());
        }

        [Fact]
        public async Task TruncatedRecordAsyncThrows()
        {
            byte[] full = [.. B.Record(Brt.RowHdr), .. B.Record(Brt.CellBool, B.CellBool(0, 0, true))];
            byte[] truncated = full[..^1];
            await using var e = Open(BlankReader(), truncated);
            await Assert.ThrowsAsync<InvalidDataException>(async () => await e.MoveNextAsync());
        }

        // --- Cancellation ---

        [Fact]
        public void MoveNextHonorsCancellation()
        {
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            byte[] sheet = [.. B.Record(Brt.RowHdr), .. B.Record(Brt.CellBool, B.CellBool(0, 0, true))];
            using var e = Open(BlankReader(), sheet, ct: cts.Token);
            Assert.Throws<OperationCanceledException>(() => e.MoveNext());
        }

        [Fact]
        public async Task MoveNextAsyncHonorsCancellation()
        {
            using var cts = new CancellationTokenSource();
            cts.Cancel();
            byte[] sheet = [.. B.Record(Brt.RowHdr), .. B.Record(Brt.CellBool, B.CellBool(0, 0, true))];
            await using var e = Open(BlankReader(), sheet, ct: cts.Token);
            await Assert.ThrowsAsync<OperationCanceledException>(async () => await e.MoveNextAsync());
        }

        // --- Buffer refill / growth across many rows ---

        private static byte[] BuildManyRowsSheet(int rowCount)
        {
            List<byte> bytes = [];
            for (int i = 0; i < rowCount; i++)
            {
                bytes.AddRange(B.Record(Brt.RowHdr));
                bytes.AddRange(B.Record(Brt.CellSt, B.CellSt(0, 0, $"row-value-{i}")));
            }
            return [.. bytes];
        }

        [Fact]
        public void ManyRowsSpanningMultipleBufferFillsEnumerateInOrderSync()
        {
            const int rowCount = 300;
            byte[] sheet = BuildManyRowsSheet(rowCount);
            // entryLength > 0 clamps the initial buffer to its 4 KB floor, forcing multiple
            // Fill/PrepareBuffer(grow) cycles since the sheet is well over 4 KB.
            using var e = Open(BlankReader(), sheet, entryLength: 1);

            for (int i = 0; i < rowCount; i++)
            {
                Assert.True(e.MoveNext());
                Assert.Equal($"row-value-{i}", e.Current[0].GetString());
            }
            Assert.False(e.MoveNext());
        }

        [Fact]
        public async Task ManyRowsSpanningMultipleBufferFillsEnumerateInOrderAsync()
        {
            const int rowCount = 300;
            byte[] sheet = BuildManyRowsSheet(rowCount);
            await using var e = Open(BlankReader(), sheet, entryLength: 1);

            for (int i = 0; i < rowCount; i++)
            {
                Assert.True(await e.MoveNextAsync());
                Assert.Equal($"row-value-{i}", e.Current[0].GetString());
            }
            Assert.False(await e.MoveNextAsync());
        }
    }
}
