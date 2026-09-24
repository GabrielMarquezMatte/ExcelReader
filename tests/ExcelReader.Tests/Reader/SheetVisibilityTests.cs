using System.IO.Compression;
using System.Text;
using ExcelReader.Core;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Xls;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Csv;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;
using ExcelReader.Tests.Reader.Xls;
using B = ExcelReader.Tests.Reader.Xlsb.Biff12Build;

namespace ExcelReader.Tests.Reader
{
    public class SheetVisibilityTests
    {
        private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string PkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";

        /// <summary>Builds an XLSX whose sheets carry the given <c>state</c> attribute text; a null entry omits it.</summary>
        private static byte[] BuildXlsx(params string?[] states)
        {
            var sheetXml = new string[states.Length];
            var relXml = new string[states.Length];
            var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                for (int i = 0; i < states.Length; i++)
                {
                    int id = i + 1;
                    string state = states[i] is null ? "" : $""" state="{states[i]}" """.TrimEnd();
                    sheetXml[i] = $"""<sheet name="S{id}" sheetId="{id}"{state} r:id="rId{id}"/>""";
                    relXml[i] = $"""<Relationship Id="rId{id}" Type="x" Target="worksheets/sheet{id}.xml"/>""";
                    Write(zip, $"xl/worksheets/sheet{id}.xml", $"""<worksheet xmlns="{Main}"><sheetData/></worksheet>""");
                }
                Write(zip, "xl/workbook.xml",
                    $"""<workbook xmlns="{Main}" xmlns:r="{Rel}"><sheets>{string.Concat(sheetXml)}</sheets></workbook>""");
                Write(zip, "xl/_rels/workbook.xml.rels",
                    $"""<Relationships xmlns="{PkgRel}">{string.Concat(relXml)}</Relationships>""");
            }
            return ms.ToArray();
        }

        private static void Write(ZipArchive zip, string path, string content)
        {
            using Stream entry = zip.CreateEntry(path).Open();
            entry.Write(Encoding.UTF8.GetBytes(content));
        }

        [Fact]
        public void Should_ReportEverySheetsState_When_ReadingXlsx()
        {
            byte[] xlsx = BuildXlsx(null, "hidden", "veryHidden");

            using var stream = new MemoryStream(xlsx, writable: false);
            using XlsxReader reader = Excel.FromXlsx(stream);

            Assert.Equal(ExcelSheetVisibility.Visible, reader.SheetVisibilityAt(0));
            Assert.Equal(ExcelSheetVisibility.Hidden, reader.SheetVisibilityAt(1));
            Assert.Equal(ExcelSheetVisibility.VeryHidden, reader.SheetVisibilityAt(2));
        }

        [Fact]
        public void Should_FollowTheCurrentSheet_When_MovingBetweenSheets()
        {
            byte[] xlsx = BuildXlsx("visible", "hidden");

            using var stream = new MemoryStream(xlsx, writable: false);
            using IExcelRowReader reader = Excel.Open(stream);

            Assert.Equal(ExcelSheetVisibility.Visible, reader.SheetVisibility);
            reader.MoveToSheet(1);
            Assert.Equal(ExcelSheetVisibility.Hidden, reader.SheetVisibility);
        }

        [Fact]
        public void Should_ReportEverySheetsState_When_ReadingXlsxFromMemory()
        {
            byte[] xlsx = BuildXlsx(null, "veryHidden");

            using XlsxReader reader = Excel.FromXlsx(xlsx.AsMemory());

            Assert.Equal(ExcelSheetVisibility.Visible, reader.SheetVisibilityAt(0));
            Assert.Equal(ExcelSheetVisibility.VeryHidden, reader.SheetVisibilityAt(1));
        }

        [Fact]
        public void Should_TreatUnknownStatesAsVisible_When_ReadingXlsx()
        {
            byte[] xlsx = BuildXlsx("", "somethingElse", "VERYHIDDEN");

            using XlsxReader reader = Excel.FromXlsx(xlsx.AsMemory());

            Assert.Equal(ExcelSheetVisibility.Visible, reader.SheetVisibilityAt(0));
            Assert.Equal(ExcelSheetVisibility.Visible, reader.SheetVisibilityAt(1));
            Assert.Equal(ExcelSheetVisibility.VeryHidden, reader.SheetVisibilityAt(2));
        }

        [Fact]
        public void Should_CarryVisibility_When_WalkingSheets()
        {
            byte[] xlsx = BuildXlsx(null, "hidden", "veryHidden");

            using IExcelRowReader reader = Excel.Open(xlsx.AsMemory());
            ExcelSheet[] sheets = [.. reader.Sheets()];

            Assert.Equal(
                [ExcelSheetVisibility.Visible, ExcelSheetVisibility.Hidden, ExcelSheetVisibility.VeryHidden],
                sheets.Select(static s => s.Visibility));
            Assert.Equal([0, 1, 2], sheets.Select(static s => s.Index));
        }

        private const int BrtBundleSh = 156;

        [Theory]
        [InlineData(0u, ExcelSheetVisibility.Visible)]
        [InlineData(1u, ExcelSheetVisibility.Hidden)]
        [InlineData(2u, ExcelSheetVisibility.VeryHidden)]
        [InlineData(7u, ExcelSheetVisibility.Visible)]
        public void Should_DecodeHsState_When_ReadingXlsbBundleSh(uint hsState, ExcelSheetVisibility expected)
        {
            byte[] workbook =
            [
                .. B.Record(BrtBundleSh, [.. B.U32(hsState), .. B.U32(0), .. B.WideString("rId1"), .. B.WideString("Plan1")]),
            ];
            byte[] rels = Encoding.UTF8.GetBytes(
                """<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.bin"/></Relationships>""");

            var sheets = XlsbWorkbook.ParseSheets(workbook, rels);

            Assert.Equal(expected, Assert.Single(sheets).Visibility);
        }

        [Fact]
        public void Should_DecodeHsState_When_ReadingXlsBoundSheet()
        {
            using MemoryStream xls = XlsWorkbookBuilder.Build(
                sheetStates: [0, 1, 2],
                sheets: [("S1", [["A"]]), ("S2", [["B"]]), ("S3", [["C"]])]);

            using XlsReader reader = Excel.FromXls(xls);

            Assert.Equal(ExcelSheetVisibility.Visible, reader.SheetVisibilityAt(0));
            Assert.Equal(ExcelSheetVisibility.Hidden, reader.SheetVisibilityAt(1));
            Assert.Equal(ExcelSheetVisibility.VeryHidden, reader.SheetVisibilityAt(2));
        }

        [Fact]
        public void Should_ReportVisible_When_ReadingCsv()
        {
            using CsvReader reader = Excel.FromCsv("a,b\n"u8.ToArray().AsMemory());

            Assert.Equal(ExcelSheetVisibility.Visible, reader.SheetVisibility);
            Assert.Equal(ExcelSheetVisibility.Visible, reader.SheetVisibilityAt(0));
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.SheetVisibilityAt(1));
        }

        private static readonly (string Name, ExcelSheetVisibility Visibility)[] _mixedSheets =
        [
            ("Visible", ExcelSheetVisibility.Visible),
            ("Hidden", ExcelSheetVisibility.Hidden),
            ("VeryHidden", ExcelSheetVisibility.VeryHidden),
        ];

        private static void WriteSheets<TWorkbook, TSheet, TRow>(
            TWorkbook workbook, params (string Name, ExcelSheetVisibility Visibility)[] sheets)
            where TWorkbook : IWorkbookWriter<TSheet>
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            foreach ((string name, ExcelSheetVisibility visibility) in sheets)
            {
                using TSheet sheet = workbook.AddSheet(name, visibility);
                using (TRow row = sheet.StartRow())
                {
                    row.Write(name);
                }
                sheet.End();
            }
            workbook.End();
        }

        private static void AssertRoundTrips(MemoryStream written)
        {
            written.Position = 0;
            using IExcelRowReader reader = Excel.Open(written);

            Assert.Equal(_mixedSheets.Length, reader.SheetCount);
            for (int i = 0; i < _mixedSheets.Length; i++)
            {
                Assert.Equal(_mixedSheets[i].Name, reader.SheetNameAt(i));
                Assert.Equal(_mixedSheets[i].Visibility, reader.SheetVisibilityAt(i));
            }
        }

        [Fact]
        public void Should_RoundTripVisibility_When_WritingXlsx()
        {
            using var ms = new MemoryStream();
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                WriteSheets<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(workbook, _mixedSheets);
            }

            AssertRoundTrips(ms);
        }

        [Fact]
        public void Should_RoundTripVisibility_When_WritingXlsb()
        {
            using var ms = new MemoryStream();
            using (XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                WriteSheets<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(workbook, _mixedSheets);
            }

            AssertRoundTrips(ms);
        }

        [Fact]
        public void Should_RoundTripVisibility_When_WritingXls()
        {
            using var ms = new MemoryStream();
            using (XlsWorkbookWriter workbook = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                WriteSheets<XlsWorkbookWriter, XlsSheetWriter, XlsRowWriter>(workbook, _mixedSheets);
            }

            AssertRoundTrips(ms);
        }

        [Fact]
        public void Should_DefaultToVisible_When_AddingASheetWithoutAVisibility()
        {
            using var ms = new MemoryStream();
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(ms, leaveOpen: true))
            {
                WriteSheets<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(workbook, ("Plain", ExcelSheetVisibility.Visible));
            }

            ms.Position = 0;
            using IExcelRowReader reader = Excel.Open(ms);
            Assert.Equal(ExcelSheetVisibility.Visible, reader.SheetVisibility);
            Assert.DoesNotContain("state=", Encoding.UTF8.GetString(WorkbookPart(ms)), StringComparison.Ordinal);
        }

        private static byte[] WorkbookPart(MemoryStream xlsx)
        {
            xlsx.Position = 0;
            using var zip = new ZipArchive(xlsx, ZipArchiveMode.Read, leaveOpen: true);
            using Stream entry = zip.GetEntry("xl/workbook.xml")!.Open();
            using var copy = new MemoryStream();
            entry.CopyTo(copy);
            return copy.ToArray();
        }

        [Theory]
        [InlineData("xlsx")]
        [InlineData("xlsb")]
        [InlineData("xls")]
        public void Should_Reject_When_EverySheetWouldBeHidden(string format)
        {
            (string, ExcelSheetVisibility)[] allHidden =
            [
                ("One", ExcelSheetVisibility.Hidden),
                ("Two", ExcelSheetVisibility.VeryHidden),
            ];

            using var ms = new MemoryStream();
            Action write = format switch
            {
                "xlsx" => () =>
                {
                    using XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
                    WriteSheets<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(wb, allHidden);
                }
                ,
                "xlsb" => () =>
                {
                    using XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true);
                    WriteSheets<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(wb, allHidden);
                }
                ,
                _ => () =>
                {
                    using XlsWorkbookWriter wb = XlsWorkbookWriter.Create(ms, leaveOpen: true);
                    WriteSheets<XlsWorkbookWriter, XlsSheetWriter, XlsRowWriter>(wb, allHidden);
                }
                ,
            };

            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(write);
            Assert.Contains("at least one visible sheet", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Should_Reject_When_TheVisibilityIsNotADefinedValue()
        {
            using var ms = new MemoryStream();
            using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(ms, leaveOpen: true);

            Assert.Throws<ArgumentOutOfRangeException>(() => workbook.AddSheet("S1", (ExcelSheetVisibility)9));
        }

        [Fact]
        public void Should_AcceptAndIgnoreVisibility_When_WritingCsv()
        {
            using var ms = new MemoryStream();
            using (CsvWorkbookWriter workbook = CsvWorkbookWriter.Create(ms, leaveOpen: true))
            {
                using CsvSheetWriter sheet = workbook.AddSheet("Ignored", ExcelSheetVisibility.VeryHidden);
                using (CsvRowWriter row = sheet.StartRow())
                {
                    row.Write("a");
                }
                sheet.End();
                workbook.End();
            }

            Assert.Equal("a\r\n", Encoding.UTF8.GetString(ms.ToArray()));
        }

        [Fact]
        public void Should_Reject_When_TheCsvVisibilityIsNotADefinedValue()
        {
            using var ms = new MemoryStream();
            using CsvWorkbookWriter workbook = CsvWorkbookWriter.Create(ms, leaveOpen: true);

            Assert.Throws<ArgumentOutOfRangeException>(() => workbook.AddSheet("S1", (ExcelSheetVisibility)9));
        }

        [Fact]
        public void Should_Reject_When_TheSheetIndexIsOutOfRange()
        {
            using XlsxReader reader = Excel.FromXlsx(BuildXlsx([null]).AsMemory());

            Assert.Throws<ArgumentOutOfRangeException>(() => reader.SheetVisibilityAt(-1));
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.SheetVisibilityAt(1));
        }
    }
}
