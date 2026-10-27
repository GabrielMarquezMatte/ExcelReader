using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xls;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Tests.Crypto;

namespace ExcelReader.Tests.Reader
{
    public sealed class StreamPositionTests : IDisposable
    {
        private const int Prefix = 16;

        public enum Kind
        {
            ExposedMemoryStream,
            ExposedMemoryStreamWithOrigin,
            OpaqueMemoryStream,
            Trickle,
            FileStream,
        }

        public static TheoryData<Kind> Kinds =>
        [
            Kind.ExposedMemoryStream, Kind.ExposedMemoryStreamWithOrigin, Kind.OpaqueMemoryStream,
            Kind.Trickle, Kind.FileStream,
        ];

        private readonly List<string> _files = [];

        public void Dispose()
        {
            foreach (string file in _files)
            {
                File.Delete(file);
            }
        }

        private static byte[] WithPrefix(byte[] workbook, int prefix = Prefix)
        {
            byte[] bytes = new byte[prefix + workbook.Length];
            Array.Fill(bytes, (byte)0x5A, 0, prefix);
            workbook.CopyTo(bytes, prefix);
            return bytes;
        }

        private Stream StreamAtPrefix(Kind kind, byte[] workbook)
        {
            byte[] bytes = WithPrefix(workbook);
            switch (kind)
            {
                case Kind.ExposedMemoryStream:
                    MemoryStream exposed = new();
                    exposed.Write(bytes);
                    exposed.Position = Prefix;
                    return exposed;
                case Kind.ExposedMemoryStreamWithOrigin:
                    byte[] shifted = WithPrefix(bytes, 5);
                    return new MemoryStream(shifted, 5, bytes.Length, writable: false, publiclyVisible: true) { Position = Prefix };
                case Kind.OpaqueMemoryStream:
                    return new MemoryStream(bytes, writable: false) { Position = Prefix };
                case Kind.Trickle:
                    return new TrickleStream(bytes) { Position = Prefix };
                default:
                    string path = Path.Combine(Path.GetTempPath(), "excelreader-position-" + Guid.NewGuid().ToString("N") + ".bin");
                    File.WriteAllBytes(path, bytes);
                    _files.Add(path);
                    var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                    file.Position = Prefix;
                    return file;
            }
        }

        private static async Task<byte[]> XlsxBytesAsync()
        {
            await using MemoryStream ms = await TypedWorkbook.BuildAsync(["Id", "Name"], [42, "alice"], [7, "bob"]);
            return ms.ToArray();
        }

        private static async Task<byte[]> XlsbBytesAsync(CancellationToken ct)
        {
            using MemoryStream ms = new();
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await using (XlsbRowWriter header = await sheet.StartRowAsync(ct))
                {
                    header.Write("Id");
                }
                await using (XlsbRowWriter row = await sheet.StartRowAsync(ct))
                {
                    row.Write(42);
                }
                await sheet.EndAsync(ct);
            }
            return ms.ToArray();
        }

        private static byte[] XlsBytes()
        {
            return File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "data", "excel-formula-strings.xls"));
        }

        private static string Dump(IExcelWorkbook workbook)
        {
            var text = new StringBuilder();
            using IExcelRowEnumerator rows = workbook.FirstSheet.GetEnumerator();
            while (rows.MoveNext())
            {
                foreach (RowCell cell in rows.Current.Cells)
                {
                    text.Append(cell.ColumnIndex).Append('=').Append(cell.Value.GetString()).Append(';');
                }
                text.Append('|');
            }
            return text.ToString();
        }

        private static string Expected(byte[] workbook)
        {
            using IExcelWorkbook reference = Excel.Open(workbook);
            return Dump(reference);
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public async Task OpenTreatsThePositionAsTheStartOfAnXlsx(Kind kind)
        {
            byte[] xlsx = await XlsxBytesAsync();
            using Stream stream = StreamAtPrefix(kind, xlsx);

            using IExcelWorkbook workbook = Excel.Open(stream);

            Assert.Equal(Expected(xlsx), Dump(workbook));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public async Task OpenAsyncTreatsThePositionAsTheStartOfAnXlsx(Kind kind)
        {
            byte[] xlsx = await XlsxBytesAsync();
            await using Stream stream = StreamAtPrefix(kind, xlsx);

            await using IExcelWorkbook workbook = await Excel.OpenAsync(stream, ct: TestContext.Current.CancellationToken);

            Assert.Equal(Expected(xlsx), Dump(workbook));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public async Task FromXlsxAndFromXlsxAsyncReadFromThePosition(Kind kind)
        {
            byte[] xlsx = await XlsxBytesAsync();
            using Stream sync = StreamAtPrefix(kind, xlsx);
            await using Stream async = StreamAtPrefix(kind, xlsx);

            using XlsxWorkbook first = Excel.FromXlsx(sync);
            await using XlsxWorkbook second = await Excel.FromXlsxAsync(async, ct: TestContext.Current.CancellationToken);

            Assert.Equal(Expected(xlsx), Dump(first));
            Assert.Equal(Expected(xlsx), Dump(second));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public async Task XlsbReadsFromThePosition(Kind kind)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] xlsb = await XlsbBytesAsync(ct);
            using Stream viaOpen = StreamAtPrefix(kind, xlsb);
            using Stream viaFrom = StreamAtPrefix(kind, xlsb);

            using IExcelWorkbook opened = Excel.Open(viaOpen);
            using XlsbWorkbook from = Excel.FromXlsb(viaFrom);

            Assert.Equal(Expected(xlsb), Dump(opened));
            Assert.Equal(Expected(xlsb), Dump(from));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public async Task XlsReadsFromThePosition(Kind kind)
        {
            byte[] xls = XlsBytes();
            using Stream viaOpen = StreamAtPrefix(kind, xls);
            using Stream viaFrom = StreamAtPrefix(kind, xls);
            await using Stream viaAsync = StreamAtPrefix(kind, xls);

            using IExcelWorkbook opened = Excel.Open(viaOpen);
            using XlsWorkbook from = Excel.FromXls(viaFrom);
            await using XlsWorkbook fromAsync = await Excel.FromXlsAsync(viaAsync, ct: TestContext.Current.CancellationToken);

            string expected = Expected(xls);
            Assert.Equal(expected, Dump(opened));
            Assert.Equal(expected, Dump(from));
            Assert.Equal(expected, Dump(fromAsync));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public async Task EncryptedWorkbooksReadFromThePosition(Kind kind)
        {
            const string name = "agile-aes256-sha512.xlsx";
            var options = new ExcelReaderOptions { Password = EncryptedFixtures.PasswordFor(name) };
            byte[] encrypted = EncryptedFixtures.Bytes(name);
            string expected = Expected(EncryptedFixtures.PlainBytes(name));
            using Stream viaOpen = StreamAtPrefix(kind, encrypted);
            using Stream viaFrom = StreamAtPrefix(kind, encrypted);
            await using Stream viaAsync = StreamAtPrefix(kind, encrypted);

            using IExcelWorkbook opened = Excel.Open(viaOpen, options: options);
            using XlsxWorkbook from = Excel.FromXlsx(viaFrom, options: options);
            await using IExcelWorkbook openedAsync = await Excel.OpenAsync(viaAsync, options: options, ct: TestContext.Current.CancellationToken);

            Assert.Equal(expected, Dump(opened));
            Assert.Equal(expected, Dump(from));
            Assert.Equal(expected, Dump(openedAsync));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public async Task DetectFileFormatStartsAtThePositionAndRestoresIt(Kind kind)
        {
            byte[] xlsx = await XlsxBytesAsync();
            byte[] xls = XlsBytes();
            using Stream zipStream = StreamAtPrefix(kind, xlsx);
            using Stream cfbStream = StreamAtPrefix(kind, xls);
            await using Stream asyncStream = StreamAtPrefix(kind, xlsx);

            Assert.Equal(ExcelFileFormat.Xlsx, Excel.DetectFileFormat(zipStream));
            Assert.Equal(Prefix, zipStream.Position);
            Assert.Equal(ExcelFileFormat.Xls, Excel.DetectFileFormat(cfbStream));
            Assert.Equal(Prefix, cfbStream.Position);
            Assert.Equal(ExcelFileFormat.Xlsx, await Excel.DetectFileFormatAsync(asyncStream, TestContext.Current.CancellationToken));
            Assert.Equal(Prefix, asyncStream.Position);
        }

        [Fact]
        public async Task AStreamLeftAtItsEndIsEmptyForEveryOpener()
        {
            byte[] xlsx = await XlsxBytesAsync();
            using MemoryStream open = new(xlsx) { Position = xlsx.Length };
            using MemoryStream from = new(xlsx) { Position = xlsx.Length };

            Assert.Throws<InvalidDataException>(() => Excel.Open(open));
            Assert.ThrowsAny<InvalidDataException>(() => Excel.FromXlsx(from));
        }
    }
}
