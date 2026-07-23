using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    public class ExcelOpenAndOleErrorTests
    {
        // --- Excel.Open format detection ---

        [Fact]
        public void OpenDetectsXlsxFromStream()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t>Hello</t></is></c></row>""");
            using var reader = Excel.Open(ms);

            XlsxReader xlsx = Assert.IsType<XlsxReader>(reader);
            using XlsxReader.Enumerator e = xlsx.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("Hello", e.Current[0].GetString());
        }

        [Fact]
        public void OpenDetectsXlsFromStream()
        {
            using MemoryStream ms = XlsWorkbookBuilder.Build(sheets: [("S1", [["A"]])]);
            using var reader = Excel.Open(ms);

            XlsReader xls = Assert.IsType<XlsReader>(reader);
            using XlsReader.Enumerator e = xls.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("A", e.Current[0].GetString());
        }

        [Fact]
        public async Task OpenAsyncDetectsBothFormatsFromStream()
        {
            await using MemoryStream xlsxBytes = WorkbookBuilder.Build(
                """<row r="1"><c r="A1"><v>1</v></c></row>""");
            await using (var xlsx = await Excel.OpenAsync(xlsxBytes, ct: TestContext.Current.CancellationToken))
            {
                Assert.IsType<XlsxReader>(xlsx);
            }

            await using MemoryStream xlsBytes = XlsWorkbookBuilder.Build(sheets: [("S1", [["A"]])]);
            await using var xls = await Excel.OpenAsync(xlsBytes, ct: TestContext.Current.CancellationToken);
            Assert.IsType<XlsReader>(xls);
        }

        [Fact]
        public void OpenDetectsFromFilePath()
        {
            string xlsxPath = WriteTemp(".xlsx", WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>"""));
            string xlsPath = WriteTemp(".xls", XlsWorkbookBuilder.Build(sheets: [("S1", [["A"]])]));
            try
            {
                using (var xlsx = Excel.Open(xlsxPath))
                {
                    Assert.IsType<XlsxReader>(xlsx);
                }
                using var xls = Excel.Open(xlsPath);
                Assert.IsType<XlsReader>(xls);
            }
            finally
            {
                File.Delete(xlsxPath);
                File.Delete(xlsPath);
            }
        }

        [Fact]
        public async Task OpenAsyncDetectsFromFilePath()
        {
            string path = WriteTemp(".xls", XlsWorkbookBuilder.Build(sheets: [("S1", [["A"]])]));
            try
            {
                await using var reader = await Excel.OpenAsync(path, ct: TestContext.Current.CancellationToken);
                Assert.IsType<XlsReader>(reader);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void OpenThrowsOnUnrecognizedSignature()
        {
            using MemoryStream ms = new([0x25, 0x50, 0x44, 0x46, 0x2D]); // "%PDF-"
            Assert.Throws<InvalidDataException>(() => Excel.Open(ms));
        }

        [Fact]
        public void OpenThrowsOnEmptyStream()
        {
            using MemoryStream ms = new();
            Assert.Throws<InvalidDataException>(() => Excel.Open(ms));
        }

        [Fact]
        public async Task OpenAsyncThrowsOnUnrecognizedSignature()
        {
            await using MemoryStream ms = new([0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07]);
            await Assert.ThrowsAsync<InvalidDataException>(async () =>
                await Excel.OpenAsync(ms, ct: TestContext.Current.CancellationToken));
        }

        [Fact]
        public void OpenThrowsOnNonSeekableStream()
        {
            using MemoryStream xls = XlsWorkbookBuilder.Build(sheets: [("S1", [["A"]])]);
            using NonSeekableStream stream = new(xls.ToArray());
            Assert.Throws<ArgumentException>(() => Excel.Open(stream));
        }

        [Fact]
        public void OpenLeavesSeekableStreamAtOriginalPositionForReader()
        {
            using MemoryStream ms = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>7</v></c></row>""");
            // A non-zero start position must be restored so the reader sees the whole stream.
            ms.Position = 0;
            using var reader = Excel.Open(ms, leaveOpen: true);
            using XlsxReader.Enumerator e = ((XlsxReader)reader).GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("7", e.Current[0].GetString());
        }

        [Fact]
        public void OpenDisposesOwnedStreamWhenSignatureUnrecognized()
        {
            // path overload opens the file with leaveOpen:false; a bad signature must not leak it.
            string path = WriteTemp(".bin", new MemoryStream([0x01, 0x02, 0x03, 0x04]));
            try
            {
                Assert.Throws<InvalidDataException>(() => Excel.Open(path));
            }
            finally
            {
                File.Delete(path); // succeeds only if Open released its FileStream handle
            }
        }

        // --- Excel.DetectFileFormat / DetectFileFormatAsync ---

        [Fact]
        public void DetectFileFormatIdentifiesXlsx()
        {
            using MemoryStream ms = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            Assert.Equal(ExcelFileFormat.Xlsx, Excel.DetectFileFormat(ms));
        }

        [Fact]
        public void DetectFileFormatIdentifiesXls()
        {
            using MemoryStream ms = XlsWorkbookBuilder.Build(sheets: [("S1", [["A"]])]);
            Assert.Equal(ExcelFileFormat.Xls, Excel.DetectFileFormat(ms));
        }

        [Fact]
        public async Task DetectFileFormatIdentifiesXlsb()
        {
            using MemoryStream ms = await BuildXlsbAsync();
            Assert.Equal(ExcelFileFormat.Xlsb, Excel.DetectFileFormat(ms));
        }

        [Fact]
        public void DetectFileFormatReturnsUnknownForUnrecognizedSignature()
        {
            using MemoryStream ms = new([0x25, 0x50, 0x44, 0x46, 0x2D]); // "%PDF-"
            Assert.Equal(ExcelFileFormat.Unknown, Excel.DetectFileFormat(ms));
        }

        [Fact]
        public void DetectFileFormatReturnsUnknownForEmptyStream()
        {
            using MemoryStream ms = new();
            Assert.Equal(ExcelFileFormat.Unknown, Excel.DetectFileFormat(ms));
        }

        [Fact]
        public void DetectFileFormatThrowsOnNullStream()
        {
            Assert.Throws<ArgumentNullException>(() => Excel.DetectFileFormat(null!));
        }

        [Fact]
        public void DetectFileFormatThrowsOnNonSeekableStream()
        {
            using MemoryStream xlsx = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            using NonSeekableStream stream = new(xlsx.ToArray());
            Assert.Throws<ArgumentException>(() => Excel.DetectFileFormat(stream));
        }

        [Fact]
        public void DetectFileFormatLeavesStreamPositionUnchanged()
        {
            using MemoryStream ms = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            ms.Position = 0;
            long before = ms.Position;

            Assert.Equal(ExcelFileFormat.Xlsx, Excel.DetectFileFormat(ms));
            Assert.Equal(before, ms.Position);

            // The stream must still be readable from the same position afterward.
            using var reader = Excel.Open(ms, leaveOpen: true);
            Assert.IsType<XlsxReader>(reader);
        }

        [Fact]
        public async Task DetectFileFormatAsyncIdentifiesXlsx()
        {
            await using MemoryStream ms = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            ExcelFileFormat format = await Excel.DetectFileFormatAsync(ms, TestContext.Current.CancellationToken);
            Assert.Equal(ExcelFileFormat.Xlsx, format);
        }

        [Fact]
        public async Task DetectFileFormatAsyncIdentifiesXls()
        {
            await using MemoryStream ms = XlsWorkbookBuilder.Build(sheets: [("S1", [["A"]])]);
            ExcelFileFormat format = await Excel.DetectFileFormatAsync(ms, TestContext.Current.CancellationToken);
            Assert.Equal(ExcelFileFormat.Xls, format);
        }

        [Fact]
        public async Task DetectFileFormatAsyncIdentifiesXlsb()
        {
            await using MemoryStream ms = await BuildXlsbAsync();
            ExcelFileFormat format = await Excel.DetectFileFormatAsync(ms, TestContext.Current.CancellationToken);
            Assert.Equal(ExcelFileFormat.Xlsb, format);
        }

        [Fact]
        public async Task DetectFileFormatAsyncReturnsUnknownForUnrecognizedSignature()
        {
            await using MemoryStream ms = new([0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07]);
            ExcelFileFormat format = await Excel.DetectFileFormatAsync(ms, TestContext.Current.CancellationToken);
            Assert.Equal(ExcelFileFormat.Unknown, format);
        }

        [Fact]
        public Task DetectFileFormatAsyncThrowsOnNullStream()
        {
            return Assert.ThrowsAsync<ArgumentNullException>(() =>
                Excel.DetectFileFormatAsync(null!, TestContext.Current.CancellationToken).AsTask());
        }

        private static async Task<MemoryStream> BuildXlsbAsync()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            MemoryStream ms = new();
            await using (XlsbWorkbookWriter wb = await XlsbWorkbookWriter.CreateAsync(ms, leaveOpen: true, ct: ct))
            {
                await wb.StartAsync(ct);
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync(ct);
                await using (XlsbRowWriter row = await sheet.StartRowAsync(ct))
                {
                    row.Write("A");
                }
                await sheet.EndAsync(ct);
                await wb.EndAsync(ct);
            }
            ms.Position = 0;
            return ms;
        }

        // --- OLE/CFB guard rails (XlsCompoundFile), reached via Excel.FromXls ---

        [Fact]
        public void CorruptOleSignatureThrows()
        {
            using MemoryStream ms = XlsWorkbookBuilder.BuildPatched(XlsWorkbookBuilder.SignatureOffset, 0x00);
            InvalidDataException ex = Assert.Throws<InvalidDataException>(() => Excel.FromXls(ms));
            Assert.Contains("OLE", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void UnsupportedOleSectorSizeThrows()
        {
            // shift 13 -> 8192-byte sectors, above the 4096 ceiling the parser accepts.
            using MemoryStream ms = XlsWorkbookBuilder.BuildPatched(
                XlsWorkbookBuilder.SectorShiftOffset, XlsWorkbookBuilder.LE16(13));
            Assert.Throws<InvalidDataException>(() => Excel.FromXls(ms));
        }

        [Fact]
        public void IncompleteDifatThrows()
        {
            // Header claims more FAT sectors than the DIFAT region actually lists.
            using MemoryStream ms = XlsWorkbookBuilder.BuildPatched(
                XlsWorkbookBuilder.FatSectorCountOffset, XlsWorkbookBuilder.LE32(9));
            Assert.Throws<InvalidDataException>(() => Excel.FromXls(ms));
        }

        [Fact]
        public void MissingWorkbookDirectoryEntryThrows()
        {
            // Rename "Workbook" -> "Xorkbook" so no Workbook/Book stream is found.
            using MemoryStream ms = XlsWorkbookBuilder.BuildPatched(
                XlsWorkbookBuilder.WorkbookEntryNameOffset, (byte)'X');
            InvalidDataException ex = Assert.Throws<InvalidDataException>(() => Excel.FromXls(ms));
            Assert.Contains("Workbook", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void NonSeekableXlsSourceIsBufferedAndRead()
        {
            byte[] bytes = XlsWorkbookBuilder.Build(sheets: [("S1", [["Name"], ["Alice"]])]).ToArray();
            using NonSeekableStream stream = new(bytes);
            using XlsReader reader = Excel.FromXls(stream, leaveOpen: false);

            using XlsReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("Name", e.Current[0].GetString());
            Assert.True(e.MoveNext());
            Assert.Equal("Alice", e.Current[0].GetString());
        }

        [Fact]
        public async Task NonSeekableXlsSourceIsBufferedAndReadAsync()
        {
            byte[] bytes = XlsWorkbookBuilder.Build(sheets: [("S1", [["Async"]])]).ToArray();
            await using NonSeekableStream stream = new(bytes);
            await using XlsReader reader = await Excel.FromXlsAsync(
                stream, leaveOpen: false, ct: TestContext.Current.CancellationToken);

            using XlsReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("Async", e.Current[0].GetString());
        }

        private static string WriteTemp(string extension, MemoryStream content)
        {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
            File.WriteAllBytes(path, content.ToArray());
            return path;
        }

    }
}
