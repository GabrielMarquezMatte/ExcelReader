using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    public class SyncAsyncParityTests
    {
        public static IEnumerable<object[]> Fixtures
        {
            get
            {
                yield return [new ParityFixture("xlsx mixed rows", BuildMixedXlsxAsync, stream => Excel.From(stream), OpenXlsxAsync)];
                yield return [new ParityFixture("xlsx refill boundary", BuildBoundaryXlsxAsync, stream => Excel.From(stream), OpenXlsxAsync)];
                yield return [new ParityFixture("xlsx many rows (mid-stream refills)", BuildManyRowsXlsxAsync, stream => Excel.From(stream), OpenXlsxAsync)];
                yield return [new ParityFixture("xls mixed rows", BuildMixedXlsAsync, stream => Excel.FromXls(stream), OpenXlsAsync)];
                yield return [new ParityFixture("xlsb mixed rows", BuildMixedXlsbAsync, stream => Excel.FromXlsb(stream), OpenXlsbAsync)];
                yield return [new ParityFixture("csv mixed rows", BuildCsvAsync, stream => Excel.FromCsv(stream), OpenCsvAsync)];
            }
        }

        [Theory]
        [MemberData(nameof(Fixtures))]
        public async Task SyncAndAsyncEnumeratorsProduceIdenticalCells(ParityFixture fixture)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] workbook = await fixture.Build(ct);

            List<CellSnapshot> sync = ReadSync(workbook, fixture.OpenSync);
            List<CellSnapshot> asyncCells = await ReadAsync(workbook, fixture.OpenAsync, ct);
            // The reader's GetAsyncEnumerator() (synchronous open, async row streaming — the 'await foreach'
            // entry point) must produce the same cells as both the sync path and the async-open path.
            List<CellSnapshot> asyncEnum = await ReadViaAsyncEnumeratorAsync(workbook, fixture.OpenSync);

            Assert.Equal(sync, asyncCells);
            Assert.Equal(sync, asyncEnum);
        }

        private static List<CellSnapshot> ReadSync(byte[] workbook, Func<Stream, IExcelRowReader> open)
        {
            using MemoryStream stream = new(workbook, writable: false);
            using IExcelRowReader reader = open(stream);
            using IExcelRowEnumerator e = reader.GetEnumerator();
            List<CellSnapshot> cells = [];
            int rowIndex = 0;
            while (e.MoveNext())
            {
                AddRow(cells, rowIndex++, e.Current);
            }
            return cells;
        }

        private static async Task<List<CellSnapshot>> ReadAsync(
            byte[] workbook,
            Func<Stream, CancellationToken, ValueTask<IExcelRowReader>> open,
            CancellationToken ct)
        {
            await using MemoryStream stream = new(workbook, writable: false);
            await using IExcelRowReader reader = await open(stream, ct);
            await using IExcelRowEnumerator e = await reader.GetAsyncEnumeratorAsync(ct);
            List<CellSnapshot> cells = [];
            int rowIndex = 0;
            while (await e.MoveNextAsync())
            {
                AddRow(cells, rowIndex++, e.Current);
            }
            return cells;
        }

        // Drives the reader's GetAsyncEnumerator() — a synchronous sheet open whose rows are then streamed
        // via MoveNextAsync (what 'await foreach' binds to). Opens the workbook synchronously; only the
        // per-row advance is awaited.
        private static async Task<List<CellSnapshot>> ReadViaAsyncEnumeratorAsync(byte[] workbook, Func<Stream, IExcelRowReader> open)
        {
            await using MemoryStream stream = new(workbook, writable: false);
            await using IExcelRowReader reader = open(stream);
            await using IExcelRowEnumerator e = reader.GetAsyncEnumerator();
            List<CellSnapshot> cells = [];
            int rowIndex = 0;
            while (await e.MoveNextAsync())
            {
                AddRow(cells, rowIndex++, e.Current);
            }
            return cells;
        }

        private static void AddRow(List<CellSnapshot> cells, int rowIndex, Row row)
        {
            cells.Add(CellSnapshot.RowMarker(rowIndex, row.ColumnCount));
            for (int column = 0; column < row.ColumnCount; column++)
            {
                Cell cell = row[column];
                bool hasDouble = cell.TryGetDouble(out double value);
                cells.Add(new CellSnapshot(
                    rowIndex,
                    column,
                    row.ColumnCount,
                    cell.Type,
                    cell.StyleIndex,
                    Convert.ToBase64String(cell.Value.ToArray()),
                    hasDouble,
                    hasDouble ? BitConverter.DoubleToInt64Bits(value) : 0));
            }
        }

        private static ValueTask<byte[]> BuildMixedXlsxAsync(CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            const string styles =
                """<styleSheet><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>""";
            const string sharedStrings =
                "<si><t>shared &lt;zero&gt;</t></si>" +
                "<si><r><t>rich</t></r><r><t> text</t></r></si>";
            const string rows =
                """<row r="1">""" +
                """<c r="A1" t="s"><v>0</v></c>""" +
                """<c r="B1" t="inlineStr"><is><t>inline &amp; entity</t></is></c>""" +
                """<c r="C1" s="1"><v>45292</v></c>""" +
                """<c r="D1" t="b"><v>1</v></c>""" +
                """<c r="E1" t="e"><v>#DIV/0!</v></c>""" +
                """<c r="F1" t="str"><v>formula text</v></c>""" +
                "</row>" +
                """<row r="2">""" +
                """<c r="A2"><v>1.25</v></c>""" +
                """<c r="D2" t="s"><v>1</v></c>""" +
                """<c r="H2" t="inlineStr"><is><t>tail</t></is></c>""" +
                "</row>";

            using MemoryStream ms = WorkbookBuilder.Build(rows, sharedStrings, styles);
            return ValueTask.FromResult(ms.ToArray());
        }

        private static ValueTask<byte[]> BuildBoundaryXlsxAsync(CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            string text = new('x', 70 * 1024);
            string rows =
                """<row r="1"><c r="A1" t="inlineStr"><is><t>""" +
                text +
                """</t></is></c><c r="C1"><v>3</v></c></row>""";

            using MemoryStream ms = WorkbookBuilder.Build(rows);
            return ValueTask.FromResult(ms.ToArray());
        }

        // Many small rows so the sheet spans several 64 KB buffer fills. Unlike the single-giant-cell
        // boundary fixture, the refills here happen deep into the sheet (when _pos is large), which is
        // where the async row-buffering slow path mishandled the compacted buffer.
        private static ValueTask<byte[]> BuildManyRowsXlsxAsync(CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            System.Text.StringBuilder sb = new(256 * 1024);
            for (int r = 1; r <= 4000; r++)
            {
                sb.Append("<row r=\"").Append(r).Append("\">")
                  .Append("<c r=\"A").Append(r).Append("\"><v>").Append(r).Append("</v></c>")
                  .Append("<c r=\"B").Append(r).Append("\" t=\"inlineStr\"><is><t>row ").Append(r).Append("</t></is></c>")
                  .Append("<c r=\"C").Append(r).Append("\"><v>").Append(r * 1.5).Append("</v></c>")
                  .Append("</row>");
            }

            using MemoryStream ms = WorkbookBuilder.Build(sb.ToString());
            return ValueTask.FromResult(ms.ToArray());
        }

        private static ValueTask<byte[]> BuildMixedXlsAsync(CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            using MemoryStream ms = XlsWorkbookBuilder.Build(
                sheets:
                [
                    ("S1",
                    [
                        ["Name", "Age", "Active", "Error", "Formula"],
                        [new XlsSharedString("Ana"), 31, true, new XlsError(0x07), new XlsFormula(12.5)],
                        [new XlsAt(2, "C"), new XlsAt(0, "A"), new XlsAt(5, new XlsRkInt(99))]
                    ])
                ]);
            return ValueTask.FromResult(ms.ToArray());
        }

        private static async ValueTask<byte[]> BuildMixedXlsbAsync(CancellationToken ct)
        {
            MemoryStream ms = new();
            await using XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: ct);
            await wb.StartAsync(ct);
            XlsbSheetWriter sheet = wb.AddSheet("S1");
            await sheet.StartAsync(ct);

            await using (XlsbRowWriter row = await sheet.StartRowAsync(ct))
            {
                row.Write("Name");
                row.Write("Age");
                row.Write("Active");
                row.Write("BirthDate");
            }
            await using (XlsbRowWriter row = await sheet.StartRowAsync(ct))
            {
                row.Write("Bia");
                row.Write(27);
                row.Write(false);
                row.Write(new DateTime(2024, 5, 6, 0, 0, 0, DateTimeKind.Unspecified));
            }
            await using (XlsbRowWriter row = await sheet.StartRowAsync(ct))
            {
                row.Write("A");
                row.Skip(2);
                row.Write("D");
            }

            await sheet.EndAsync(ct);
            await wb.EndAsync(ct);
            return ms.ToArray();
        }

        private static ValueTask<byte[]> BuildCsvAsync(CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            const string csv =
                "Name,Age,Active\n" +
                "Ana,31,true\n" +
                "\"Bia, Jr.\",27,false\n" +   // quoted field with an embedded comma
                "Cid,,\n";                     // trailing empty fields
            return ValueTask.FromResult(System.Text.Encoding.UTF8.GetBytes(csv));
        }

        private static async ValueTask<IExcelRowReader> OpenCsvAsync(Stream stream, CancellationToken ct)
        {
            return await Excel.FromCsvAsync(stream, ct: ct);
        }

        private static async ValueTask<IExcelRowReader> OpenXlsxAsync(Stream stream, CancellationToken ct)
        {
            return await Excel.FromAsync(stream, ct: ct);
        }

        private static async ValueTask<IExcelRowReader> OpenXlsAsync(Stream stream, CancellationToken ct)
        {
            return await Excel.FromXlsAsync(stream, ct: ct);
        }

        private static async ValueTask<IExcelRowReader> OpenXlsbAsync(Stream stream, CancellationToken ct)
        {
            return await Excel.FromXlsbAsync(stream, ct: ct);
        }

        public sealed record ParityFixture(
            string Name,
            Func<CancellationToken, ValueTask<byte[]>> Build,
            Func<Stream, IExcelRowReader> OpenSync,
            Func<Stream, CancellationToken, ValueTask<IExcelRowReader>> OpenAsync)
        {
            public override string ToString()
            {
                return Name;
            }
        }

        private readonly record struct CellSnapshot(
            int Row,
            int Column,
            int ColumnCount,
            CellType Type,
            int StyleIndex,
            string ValueBase64,
            bool HasDouble,
            long DoubleBits)
        {
            public static CellSnapshot RowMarker(int row, int columnCount)
            {
                return new CellSnapshot(row, -1, columnCount, CellType.Empty, 0, string.Empty, false, 0);
            }
        }
    }
}
