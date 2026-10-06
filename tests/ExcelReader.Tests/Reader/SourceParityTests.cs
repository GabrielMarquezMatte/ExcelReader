using System.Globalization;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Reader
{
    public sealed class SourceParityTests : IDisposable
    {
        public enum Kind
        {
            Memory,
            ExposedMemoryStream,
            OpaqueMemoryStream,
            Trickle,
            FileStream,
            FilePath,
            NonSeekable,
        }

        private readonly List<string> _paths = [];

        public void Dispose()
        {
            foreach (string path in _paths)
            {
                File.Delete(path);
            }
        }

        public static TheoryData<Kind> Kinds =>
        [
            Kind.Memory, Kind.ExposedMemoryStream, Kind.OpaqueMemoryStream, Kind.Trickle,
            Kind.FileStream, Kind.FilePath, Kind.NonSeekable,
        ];

        private static byte[] BuildXlsx()
        {
            using MemoryStream buffer = new();
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(buffer, leaveOpen: true, new XlsxWriterOptions { UseSharedStrings = true }))
            {
                for (int s = 0; s < 3; s++)
                {
                    XlsxSheetWriter sheet = workbook.AddSheet("sheet" + s);
                    for (int r = 0; r < 400; r++)
                    {
                        using XlsxRowWriter row = sheet.StartRow();
                        row.Write("name-" + ((s * 1000) + (r % 50)).ToString(CultureInfo.InvariantCulture));
                        row.Write((s * 100000) + r);
                        row.Write(r % 2 == 0);
                    }
                    sheet.End();
                    sheet.Dispose();
                }
                workbook.End();
            }
            return buffer.ToArray();
        }

        private string WriteTemp(byte[] bytes, string extension)
        {
            string path = Path.Combine(Path.GetTempPath(), "excelreader-parity-" + Guid.NewGuid().ToString("N") + extension);
            File.WriteAllBytes(path, bytes);
            _paths.Add(path);
            return path;
        }

        private XlsxReader OpenXlsx(Kind kind, byte[] bytes)
        {
            switch (kind)
            {
                case Kind.Memory:
                    return Excel.FromXlsx(bytes);
                case Kind.ExposedMemoryStream:
                    MemoryStream exposed = new();
                    exposed.Write(bytes);
                    exposed.Position = 0;
                    return Excel.FromXlsx(exposed, leaveOpen: false);
                case Kind.OpaqueMemoryStream:
                    return Excel.FromXlsx(new MemoryStream(bytes, writable: false), leaveOpen: false);
                case Kind.Trickle:
                    return Excel.FromXlsx(new TrickleStream(bytes), leaveOpen: false);
                case Kind.FileStream:
                    return Excel.FromXlsx(File.OpenRead(WriteTemp(bytes, ".xlsx")), leaveOpen: false);
                case Kind.FilePath:
                    return Excel.FromXlsxFile(WriteTemp(bytes, ".xlsx"));
                default:
                    return Excel.FromXlsx(new NonSeekableStream(bytes), leaveOpen: false);
            }
        }

        private static List<string> Dump(IExcelRowReader reader)
        {
            List<string> lines = [];
            for (int s = 0; s < reader.SheetCount; s++)
            {
                reader.MoveToSheet(s);
                using IExcelRowEnumerator e = reader.GetEnumerator();
                while (e.MoveNext())
                {
                    lines.Add(Line(s, e.Current));
                }
            }
            return lines;
        }

        private static async Task<List<string>> DumpAsync(IExcelRowReader reader)
        {
            List<string> lines = [];
            for (int s = 0; s < reader.SheetCount; s++)
            {
                reader.MoveToSheet(s);
                await using IExcelRowEnumerator e = reader.GetAsyncEnumerator(TestContext.Current.CancellationToken);
                while (await e.MoveNextAsync())
                {
                    lines.Add(Line(s, e.Current));
                }
            }
            return lines;
        }

        private static string Line(int sheet, Row row)
        {
            string[] cells = new string[row.ColumnCount];
            for (int c = 0; c < cells.Length; c++)
            {
                cells[c] = row[c].GetString();
            }
            return sheet.ToString(CultureInfo.InvariantCulture) + "|" + string.Join("|", cells);
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void Xlsx_Reads_The_Same_Rows_From_Every_Source(Kind kind)
        {
            byte[] bytes = BuildXlsx();
            using IExcelRowReader expected = Excel.FromXlsx(bytes);
            using IExcelRowReader actual = OpenXlsx(kind, bytes);
            Assert.Equal(Dump(expected), Dump(actual));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public async Task Xlsx_Reads_The_Same_Rows_Asynchronously_From_Every_Source(Kind kind)
        {
            byte[] bytes = BuildXlsx();
            using IExcelRowReader expected = Excel.FromXlsx(bytes);
            using IExcelRowReader actual = OpenXlsx(kind, bytes);
            Assert.Equal(Dump(expected), await DumpAsync(actual));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void Xlsx_Interleaved_Enumerators_Keep_Their_Own_Position(Kind kind)
        {
            byte[] bytes = BuildXlsx();
            using IExcelRowReader reader = OpenXlsx(kind, bytes);
            reader.MoveToSheet(0);
            using IExcelRowEnumerator first = reader.GetEnumerator();
            reader.MoveToSheet(1);
            using IExcelRowEnumerator second = reader.GetEnumerator();
            for (int r = 0; r < 400; r++)
            {
                Assert.True(first.MoveNext());
                Assert.True(second.MoveNext());
                Assert.Equal("name-" + (r % 50).ToString(CultureInfo.InvariantCulture), first.Current[0].GetString());
                Assert.Equal("name-" + (1000 + (r % 50)).ToString(CultureInfo.InvariantCulture), second.Current[0].GetString());
            }
            Assert.False(first.MoveNext());
            Assert.False(second.MoveNext());
        }

        [Fact]
        public void Xlsx_File_Can_Be_Deleted_After_Dispose()
        {
            string path = WriteTemp(BuildXlsx(), ".xlsx");
            using (IExcelRowReader reader = Excel.FromXlsxFile(path))
            {
                Assert.NotEmpty(Dump(reader));
            }
            File.Delete(path);
            Assert.False(File.Exists(path));
        }

        [Fact]
        public void Xlsx_Leaves_A_Borrowed_Stream_Usable()
        {
            byte[] bytes = BuildXlsx();
            using TrickleStream stream = new(bytes);
            using (IExcelRowReader reader = Excel.FromXlsx(stream, leaveOpen: true))
            {
                Assert.NotEmpty(Dump(reader));
            }
            Assert.True(stream.CanRead);
        }

        [Fact]
        public void Xlsx_Cut_Short_Is_Invalid_Data()
        {
            byte[] bytes = BuildXlsx();
            string path = WriteTemp(bytes[..(bytes.Length / 2)], ".xlsx");
            Assert.Throws<InvalidDataException>(() => Excel.FromXlsxFile(path));
        }

        private static byte[] BuildXlsb()
        {
            using MemoryStream buffer = new();
            using (XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(buffer, leaveOpen: true, new XlsbWriterOptions { UseSharedStrings = true }))
            {
                for (int s = 0; s < 3; s++)
                {
                    XlsbSheetWriter sheet = workbook.AddSheet("sheet" + s);
                    for (int r = 0; r < 400; r++)
                    {
                        using XlsbRowWriter row = sheet.StartRow();
                        row.Write("name-" + ((s * 1000) + (r % 50)).ToString(CultureInfo.InvariantCulture));
                        row.Write((s * 100000) + r);
                        row.Write(r % 2 == 0);
                    }
                    sheet.End();
                    sheet.Dispose();
                }
                workbook.End();
            }
            return buffer.ToArray();
        }

        private XlsbReader OpenXlsb(Kind kind, byte[] bytes)
        {
            switch (kind)
            {
                case Kind.Memory:
                    return Excel.FromXlsb(bytes);
                case Kind.ExposedMemoryStream:
                    MemoryStream exposed = new();
                    exposed.Write(bytes);
                    exposed.Position = 0;
                    return Excel.FromXlsb(exposed, leaveOpen: false);
                case Kind.OpaqueMemoryStream:
                    return Excel.FromXlsb(new MemoryStream(bytes, writable: false), leaveOpen: false);
                case Kind.Trickle:
                    return Excel.FromXlsb(new TrickleStream(bytes), leaveOpen: false);
                case Kind.FileStream:
                    return Excel.FromXlsb(File.OpenRead(WriteTemp(bytes, ".xlsb")), leaveOpen: false);
                case Kind.FilePath:
                    return Excel.FromXlsbFile(WriteTemp(bytes, ".xlsb"));
                default:
                    return Excel.FromXlsb(new NonSeekableStream(bytes), leaveOpen: false);
            }
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void Xlsb_Reads_The_Same_Rows_From_Every_Source(Kind kind)
        {
            byte[] bytes = BuildXlsb();
            using IExcelRowReader expected = Excel.FromXlsb(bytes);
            using IExcelRowReader actual = OpenXlsb(kind, bytes);
            Assert.Equal(Dump(expected), Dump(actual));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public async Task Xlsb_Reads_The_Same_Rows_Asynchronously_From_Every_Source(Kind kind)
        {
            byte[] bytes = BuildXlsb();
            using IExcelRowReader expected = Excel.FromXlsb(bytes);
            using IExcelRowReader actual = OpenXlsb(kind, bytes);
            Assert.Equal(Dump(expected), await DumpAsync(actual));
        }

        [Fact]
        public void Xlsb_File_Can_Be_Deleted_After_Dispose()
        {
            string path = WriteTemp(BuildXlsb(), ".xlsb");
            using (IExcelRowReader reader = Excel.FromXlsbFile(path))
            {
                Assert.NotEmpty(Dump(reader));
            }
            File.Delete(path);
            Assert.False(File.Exists(path));
        }

        [Fact]
        public void Xlsb_Cut_Short_Is_Invalid_Data()
        {
            byte[] bytes = BuildXlsb();
            string path = WriteTemp(bytes[..(bytes.Length / 2)], ".xlsb");
            Assert.Throws<InvalidDataException>(() => Excel.FromXlsbFile(path));
        }
    }
}
