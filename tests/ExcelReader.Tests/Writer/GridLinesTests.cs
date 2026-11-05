using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Csv;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer
{
    public sealed class GridLinesTests
    {
        private const int XlsbDisplayGridLines = 0x04;
        private const int XlsDisplayGridLines = 0x02;

        private static byte[] Write<TWorkbook, TSheet, TRow>(Func<Stream, TWorkbook> create, bool? showGridLines, int rows = 1)
            where TWorkbook : IWorkbookWriter<TSheet>
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            using MemoryStream stream = new();
            using (TWorkbook workbook = create(stream))
            {
                using TSheet sheet = workbook.AddSheet("Data");
                if (showGridLines is bool show)
                {
                    sheet.SetShowGridLines(show);
                }
                for (int i = 0; i < rows; i++)
                {
                    using TRow row = sheet.StartRow();
                    row.Write("v");
                }
                sheet.End();
                workbook.End();
            }
            return stream.ToArray();
        }

        private static byte[] WriteXlsx(bool? showGridLines)
        {
            return Write<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(static s => XlsxWorkbookWriter.Create(s, leaveOpen: true), showGridLines);
        }

        private static byte[] WriteXlsb(bool? showGridLines)
        {
            return Write<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(static s => XlsbWorkbookWriter.Create(s, leaveOpen: true), showGridLines);
        }

        private static byte[] WriteXls(bool? showGridLines, int rows = 1)
        {
            return Write<XlsWorkbookWriter, XlsSheetWriter, XlsRowWriter>(static s => XlsWorkbookWriter.Create(s, leaveOpen: true), showGridLines, rows);
        }

        private static byte[] Entry(byte[] package, string name)
        {
            using ZipArchive zip = new(new MemoryStream(package), ZipArchiveMode.Read);
            using Stream entry = zip.GetEntry(name)!.Open();
            using MemoryStream copy = new();
            entry.CopyTo(copy);
            return copy.ToArray();
        }

        private static int ReadVarint(byte[] data, ref int position, int maxBytes)
        {
            int value = 0;
            for (int i = 0; i < maxBytes; i++)
            {
                byte b = data[position++];
                value |= (b & 0x7F) << (7 * i);
                if ((b & 0x80) == 0)
                {
                    break;
                }
            }
            return value;
        }

        private static int XlsbWorksheetViewFlags(byte[] package)
        {
            byte[] sheet = Entry(package, "xl/worksheets/sheet1.bin");
            int position = 0;
            while (position < sheet.Length)
            {
                int id = ReadVarint(sheet, ref position, 2);
                int length = ReadVarint(sheet, ref position, 4);
                if (id == Brt.BeginWsView)
                {
                    return BinaryPrimitives.ReadUInt16LittleEndian(sheet.AsSpan(position));
                }
                position += length;
            }
            throw new InvalidDataException("no BrtBeginWsView record.");
        }

        // WINDOW2 (0x023E) with its 18-byte worksheet payload; small files keep each record inside one sector.
        private static List<int> XlsWindow2Flags(byte[] file)
        {
            ReadOnlySpan<byte> header = [0x3E, 0x02, 0x12, 0x00];
            List<int> flags = [];
            ReadOnlySpan<byte> rest = file;
            int offset = rest.IndexOf(header);
            while (offset >= 0)
            {
                flags.Add(BinaryPrimitives.ReadUInt16LittleEndian(rest[(offset + 4)..]));
                rest = rest[(offset + header.Length)..];
                offset = rest.IndexOf(header);
            }
            return flags;
        }

        [Fact]
        public void Should_WriteASheetViewWithoutGridLines_When_XlsxHidesThem()
        {
            byte[] package = WriteXlsx(showGridLines: false);

            string sheet = Encoding.UTF8.GetString(Entry(package, "xl/worksheets/sheet1.xml"));
            int view = sheet.IndexOf("<sheetViews><sheetView showGridLines=\"0\" workbookViewId=\"0\"/></sheetViews>", StringComparison.Ordinal);
            Assert.True(view >= 0);
            Assert.True(view < sheet.IndexOf("<sheetData>", StringComparison.Ordinal));

            using IExcelWorkbook reader = Excel.FromXlsx(new MemoryStream(package));
            using IExcelRowEnumerator rows = reader.FirstSheet.GetEnumerator();
            Assert.True(rows.MoveNext());
            Assert.Equal("v", rows.Current[0].GetString());
        }

        [Theory]
        [InlineData(null)]
        [InlineData(true)]
        public void Should_LeaveTheSheetViewOut_When_XlsxShowsGridLines(bool? showGridLines)
        {
            string sheet = Encoding.UTF8.GetString(Entry(WriteXlsx(showGridLines), "xl/worksheets/sheet1.xml"));

            Assert.DoesNotContain("sheetView", sheet, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(null, true)]
        [InlineData(true, true)]
        [InlineData(false, false)]
        public void Should_SetTheDisplayGridLinesFlag_When_WritingXlsb(bool? showGridLines, bool expected)
        {
            int flags = XlsbWorksheetViewFlags(WriteXlsb(showGridLines));

            Assert.Equal(expected, (flags & XlsbDisplayGridLines) != 0);
            Assert.Equal(0x0398, flags & ~XlsbDisplayGridLines);
        }

        [Theory]
        [InlineData(null, true)]
        [InlineData(true, true)]
        [InlineData(false, false)]
        public void Should_SetTheDisplayGridLinesFlag_When_WritingXls(bool? showGridLines, bool expected)
        {
            int flags = Assert.Single(XlsWindow2Flags(WriteXls(showGridLines)));

            Assert.Equal(expected, (flags & XlsDisplayGridLines) != 0);
        }

        [Fact]
        public void Should_HideGridLinesOnContinuationSheets_When_AnXlsSheetOverflows()
        {
            List<int> flags = XlsWindow2Flags(WriteXls(showGridLines: false, rows: 65_537));

            Assert.Equal(2, flags.Count);
            Assert.All(flags, static f => Assert.Equal(0, f & XlsDisplayGridLines));
        }

        [Fact]
        public void Should_Throw_When_GridLinesAreSetAfterTheFirstRow()
        {
            using MemoryStream stream = new();
            using XlsxWorkbookWriter xlsx = XlsxWorkbookWriter.Create(stream, leaveOpen: true);
            using XlsxSheetWriter xlsxSheet = xlsx.AddSheet("Data");
            xlsxSheet.StartRow().Dispose();
            using XlsbWorkbookWriter xlsb = XlsbWorkbookWriter.Create(new MemoryStream(), leaveOpen: false);
            using XlsbSheetWriter xlsbSheet = xlsb.AddSheet("Data");
            xlsbSheet.StartRow().Dispose();
            using XlsWorkbookWriter xls = XlsWorkbookWriter.Create(new MemoryStream(), leaveOpen: false);
            using XlsSheetWriter xlsSheet = xls.AddSheet("Data");
            xlsSheet.StartRow().Dispose();

            Assert.Throws<InvalidOperationException>(() => xlsxSheet.SetShowGridLines(false));
            Assert.Throws<InvalidOperationException>(() => xlsbSheet.SetShowGridLines(false));
            Assert.Throws<InvalidOperationException>(() => xlsSheet.SetShowGridLines(false));
        }

        [Fact]
        public void Should_IgnoreGridLines_When_WritingCsv()
        {
            byte[] csv = Write<CsvWorkbookWriter, CsvSheetWriter, CsvRowWriter>(static s => CsvWorkbookWriter.Create(s, leaveOpen: true), showGridLines: false);

            Assert.Equal("v\r\n", Encoding.UTF8.GetString(csv));
        }
    }
}
