using System.Globalization;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xls;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer.Xls;
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

        private XlsxWorkbook OpenXlsx(Kind kind, byte[] bytes)
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

        private static List<string> Dump(IExcelWorkbook reader)
        {
            List<string> lines = [];
            for (int s = 0; s < reader.SheetCount; s++)
            {
                using IExcelRowEnumerator e = reader.SheetAt(s).GetEnumerator();
                while (e.MoveNext())
                {
                    lines.Add(Line(s, e.Current));
                }
            }
            return lines;
        }

        private static async Task<List<string>> DumpAsync(IExcelWorkbook reader)
        {
            List<string> lines = [];
            for (int s = 0; s < reader.SheetCount; s++)
            {
                await using IExcelRowEnumerator e = reader.SheetAt(s).GetAsyncEnumerator(TestContext.Current.CancellationToken);
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
            using IExcelWorkbook expected = Excel.FromXlsx(bytes);
            using IExcelWorkbook actual = OpenXlsx(kind, bytes);
            Assert.Equal(Dump(expected), Dump(actual));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public async Task Xlsx_Reads_The_Same_Rows_Asynchronously_From_Every_Source(Kind kind)
        {
            byte[] bytes = BuildXlsx();
            using IExcelWorkbook expected = Excel.FromXlsx(bytes);
            using IExcelWorkbook actual = OpenXlsx(kind, bytes);
            Assert.Equal(Dump(expected), await DumpAsync(actual));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void Xlsx_Interleaved_Enumerators_Keep_Their_Own_Position(Kind kind)
        {
            byte[] bytes = BuildXlsx();
            using IExcelWorkbook reader = OpenXlsx(kind, bytes);
            using IExcelRowEnumerator first = reader.SheetAt(0).GetEnumerator();
            using IExcelRowEnumerator second = reader.SheetAt(1).GetEnumerator();
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
            using (IExcelWorkbook reader = Excel.FromXlsxFile(path))
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
            using (IExcelWorkbook reader = Excel.FromXlsx(stream, leaveOpen: true))
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

        private static byte[] BuildXls()
        {
            using MemoryStream buffer = new();
            using (XlsWorkbookWriter workbook = XlsWorkbookWriter.Create(buffer, leaveOpen: true))
            {
                for (int s = 0; s < 3; s++)
                {
                    XlsSheetWriter sheet = workbook.AddSheet("sheet" + s);
                    for (int r = 0; r < 400; r++)
                    {
                        using XlsRowWriter row = sheet.StartRow();
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

        private XlsWorkbook OpenXls(Kind kind, byte[] bytes)
        {
            switch (kind)
            {
                case Kind.Memory:
                    return Excel.FromXls(bytes);
                case Kind.ExposedMemoryStream:
                    MemoryStream exposed = new();
                    exposed.Write(bytes);
                    exposed.Position = 0;
                    return Excel.FromXls(exposed, leaveOpen: false);
                case Kind.OpaqueMemoryStream:
                    return Excel.FromXls(new MemoryStream(bytes, writable: false), leaveOpen: false);
                case Kind.Trickle:
                    return Excel.FromXls(new TrickleStream(bytes), leaveOpen: false);
                case Kind.FileStream:
                    return Excel.FromXls(File.OpenRead(WriteTemp(bytes, ".xls")), leaveOpen: false);
                case Kind.FilePath:
                    return Excel.FromXlsFile(WriteTemp(bytes, ".xls"));
                default:
                    return Excel.FromXls(new NonSeekableStream(bytes), leaveOpen: false);
            }
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void Xls_Reads_The_Same_Rows_From_Every_Source(Kind kind)
        {
            byte[] bytes = BuildXls();
            using IExcelWorkbook expected = Excel.FromXls(bytes);
            using IExcelWorkbook actual = OpenXls(kind, bytes);
            Assert.Equal(Dump(expected), Dump(actual));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void Xls_Interleaved_Enumerators_Keep_Their_Own_Position(Kind kind)
        {
            byte[] bytes = BuildXls();
            using IExcelWorkbook reader = OpenXls(kind, bytes);
            using IExcelRowEnumerator first = reader.SheetAt(0).GetEnumerator();
            using IExcelRowEnumerator second = reader.SheetAt(1).GetEnumerator();
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
        public void Xls_File_Can_Be_Deleted_After_Dispose()
        {
            string path = WriteTemp(BuildXls(), ".xls");
            using (IExcelWorkbook reader = Excel.FromXlsFile(path))
            {
                Assert.NotEmpty(Dump(reader));
            }
            File.Delete(path);
            Assert.False(File.Exists(path));
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

        private XlsbWorkbook OpenXlsb(Kind kind, byte[] bytes)
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
            using IExcelWorkbook expected = Excel.FromXlsb(bytes);
            using IExcelWorkbook actual = OpenXlsb(kind, bytes);
            Assert.Equal(Dump(expected), Dump(actual));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public async Task Xlsb_Reads_The_Same_Rows_Asynchronously_From_Every_Source(Kind kind)
        {
            byte[] bytes = BuildXlsb();
            using IExcelWorkbook expected = Excel.FromXlsb(bytes);
            using IExcelWorkbook actual = OpenXlsb(kind, bytes);
            Assert.Equal(Dump(expected), await DumpAsync(actual));
        }

        [Fact]
        public void Xlsb_File_Can_Be_Deleted_After_Dispose()
        {
            string path = WriteTemp(BuildXlsb(), ".xlsb");
            using (IExcelWorkbook reader = Excel.FromXlsbFile(path))
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

        public static TheoryData<Kind> SeekableStreamKinds =>
        [
            Kind.ExposedMemoryStream, Kind.OpaqueMemoryStream, Kind.Trickle, Kind.FileStream,
        ];

        private Stream OpenRawStream(Kind kind, byte[] bytes)
        {
            switch (kind)
            {
                case Kind.ExposedMemoryStream:
                    MemoryStream exposed = new();
                    exposed.Write(bytes);
                    exposed.Position = 0;
                    return exposed;
                case Kind.OpaqueMemoryStream:
                    return new MemoryStream(bytes, writable: false);
                case Kind.Trickle:
                    return new TrickleStream(bytes);
                default:
                    return File.OpenRead(WriteTemp(bytes, ".bin"));
            }
        }

        [Theory]
        [MemberData(nameof(SeekableStreamKinds))]
        public void DetectFileFormat_Leaves_The_Stream_Open_At_Its_Position(Kind kind)
        {
            using Stream xlsx = OpenRawStream(kind, BuildXlsx());
            Assert.Equal(ExcelFileFormat.Xlsx, Excel.DetectFileFormat(xlsx));
            Assert.Equal(0, xlsx.Position);
            Assert.True(xlsx.CanRead);

            using Stream xlsb = OpenRawStream(kind, BuildXlsb());
            Assert.Equal(ExcelFileFormat.Xlsb, Excel.DetectFileFormat(xlsb));
            Assert.Equal(0, xlsb.Position);
            Assert.True(xlsb.CanRead);
        }

        [Theory]
        [MemberData(nameof(SeekableStreamKinds))]
        public async Task DetectFileFormatAsync_Leaves_The_Stream_Open_At_Its_Position(Kind kind)
        {
            using Stream xlsx = OpenRawStream(kind, BuildXlsx());
            Assert.Equal(ExcelFileFormat.Xlsx, await Excel.DetectFileFormatAsync(xlsx, TestContext.Current.CancellationToken));
            Assert.Equal(0, xlsx.Position);
            Assert.True(xlsx.CanRead);
        }

        [Theory]
        [MemberData(nameof(SeekableStreamKinds))]
        public void Open_Detects_And_Reads_Both_Zip_Formats(Kind kind)
        {
            byte[] xlsx = BuildXlsx();
            using IExcelWorkbook expectedXlsx = Excel.FromXlsx(xlsx);
            using IExcelWorkbook actualXlsx = Excel.Open(OpenRawStream(kind, xlsx), leaveOpen: false);
            Assert.Equal(Dump(expectedXlsx), Dump(actualXlsx));

            byte[] xlsb = BuildXlsb();
            using IExcelWorkbook expectedXlsb = Excel.FromXlsb(xlsb);
            using IExcelWorkbook actualXlsb = Excel.Open(OpenRawStream(kind, xlsb), leaveOpen: false);
            Assert.Equal(Dump(expectedXlsb), Dump(actualXlsb));
        }

        [Theory]
        [MemberData(nameof(SeekableStreamKinds))]
        public async Task OpenAsync_Detects_And_Reads_Both_Zip_Formats(Kind kind)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] xlsx = BuildXlsx();
            using IExcelWorkbook expectedXlsx = Excel.FromXlsx(xlsx);
            using IExcelWorkbook actualXlsx = await Excel.OpenAsync(OpenRawStream(kind, xlsx), leaveOpen: false, ct: ct);
            Assert.Equal(Dump(expectedXlsx), await DumpAsync(actualXlsx));

            byte[] xlsb = BuildXlsb();
            using IExcelWorkbook expectedXlsb = Excel.FromXlsb(xlsb);
            using IExcelWorkbook actualXlsb = await Excel.OpenAsync(OpenRawStream(kind, xlsb), leaveOpen: false, ct: ct);
            Assert.Equal(Dump(expectedXlsb), await DumpAsync(actualXlsb));
        }

        private static byte[] CorruptZip()
        {
            byte[] bytes = new byte[200];
            bytes.AsSpan().Fill(0x41);
            bytes[0] = 0x50;
            bytes[1] = 0x4B;
            bytes[2] = 0x03;
            bytes[3] = 0x04;
            return bytes;
        }

        private static byte[] ZipWithoutWorkbook()
        {
            using MemoryStream buffer = new();
            using (System.IO.Compression.ZipArchive zip = new(buffer, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            {
                using StreamWriter writer = new(zip.CreateEntry("hello.txt").Open());
                writer.Write("hello");
            }
            return buffer.ToArray();
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public async Task A_Corrupt_Zip_Opened_Through_Every_Entry_Point_Honours_LeaveOpen(bool leaveOpen)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = CorruptZip();
            TrickleStream open = new(bytes);
            Assert.Throws<InvalidDataException>(() => Excel.Open(open, leaveOpen));
            Assert.Equal(leaveOpen, open.CanRead);

            TrickleStream openAsync = new(bytes);
            await Assert.ThrowsAsync<InvalidDataException>(async () => await Excel.OpenAsync(openAsync, leaveOpen, ct: ct));
            Assert.Equal(leaveOpen, openAsync.CanRead);

            TrickleStream xlsx = new(bytes);
            Assert.Throws<InvalidDataException>(() => Excel.FromXlsx(xlsx, leaveOpen));
            Assert.Equal(leaveOpen, xlsx.CanRead);

            TrickleStream xlsb = new(bytes);
            Assert.Throws<InvalidDataException>(() => Excel.FromXlsb(xlsb, leaveOpen));
            Assert.Equal(leaveOpen, xlsb.CanRead);
        }

        [Fact]
        public void A_Zip_Without_A_Workbook_Closes_The_Stream()
        {
            TrickleStream stream = new(ZipWithoutWorkbook());
            Assert.Throws<InvalidDataException>(() => Excel.FromXlsx(stream, leaveOpen: false));
            Assert.False(stream.CanRead);
        }

        [Fact]
        public void A_Failed_Detection_Restores_The_Stream_Position()
        {
            TrickleStream stream = new(CorruptZip());
            Assert.Throws<InvalidDataException>(() => Excel.DetectFileFormat(stream));
            Assert.Equal(0, stream.Position);
            Assert.True(stream.CanRead);
        }

        [Fact]
        public async Task A_Failed_Asynchronous_Detection_Restores_The_Stream_Position()
        {
            TrickleStream stream = new(CorruptZip());
            await Assert.ThrowsAsync<InvalidDataException>(async () => await Excel.DetectFileFormatAsync(stream, TestContext.Current.CancellationToken));
            Assert.Equal(0, stream.Position);
            Assert.True(stream.CanRead);
        }
    }
}
