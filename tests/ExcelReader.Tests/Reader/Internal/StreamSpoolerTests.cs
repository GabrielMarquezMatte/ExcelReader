using System.Globalization;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Reader.Internal
{
    public sealed class StreamSpoolerTests
    {
        private const int First = 16;
        private const int Max = 64;

        private static byte[] Bytes(int length)
        {
            byte[] bytes = new byte[length];
            System.Security.Cryptography.RandomNumberGenerator.Fill(bytes);
            return bytes;
        }

        public static TheoryData<int> Lengths => [0, 1, First - 1, First, First + 1, First + Max, First + Max + 1, 1000];

        [Theory]
        [MemberData(nameof(Lengths))]
        public void Spool_Returns_Exact_Bytes(int length)
        {
            byte[] bytes = Bytes(length);
            Assert.Equal(bytes, StreamSpooler.Spool(new NonSeekableStream(bytes), First, Max));
        }

        [Theory]
        [MemberData(nameof(Lengths))]
        public async Task SpoolAsync_Returns_Exact_Bytes(int length)
        {
            byte[] bytes = Bytes(length);
            Assert.Equal(bytes, await StreamSpooler.SpoolAsync(new TrickleStream(bytes), TestContext.Current.CancellationToken, First, Max));
        }

        [Fact]
        public void Spool_Propagates_Read_Failure()
        {
            NonSeekableStream stream = new(Bytes(100)) { OnRead = () => throw new IOException("boom") };
            Assert.Throws<IOException>(() => StreamSpooler.Spool(stream, First, Max));
        }

        [Fact]
        public async Task SpoolAsync_Throws_When_Cancelled()
        {
            using CancellationTokenSource cts = new();
            await cts.CancelAsync();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
                await StreamSpooler.SpoolAsync(new NonSeekableStream(Bytes(100)), cts.Token, First, Max));
        }

        private static List<string> Dump(IExcelWorkbook reader)
        {
            List<string> lines = [];
            for (int s = 0; s < reader.SheetCount; s++)
            {
                using IExcelRowEnumerator e = reader.SheetAt(s).GetEnumerator();
                while (e.MoveNext())
                {
                    lines.Add(string.Join("|", Enumerable.Range(0, e.Current.ColumnCount).Select(c => e.Current[c].GetString())));
                }
            }
            return lines;
        }

        private static string Text(int r)
        {
            return Guid.NewGuid().ToString("N") + r.ToString(CultureInfo.InvariantCulture);
        }

        private static byte[] BuildXlsx()
        {
            using MemoryStream buffer = new();
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(buffer, leaveOpen: true, new XlsxWriterOptions()))
            {
                XlsxSheetWriter sheet = workbook.AddSheet("s");
                for (int r = 0; r < 12000; r++)
                {
                    using XlsxRowWriter row = sheet.StartRow();
                    row.Write(Text(r));
                    row.Write(r);
                }
                sheet.End();
                sheet.Dispose();
                workbook.End();
            }
            return buffer.ToArray();
        }

        private static byte[] BuildXlsb()
        {
            using MemoryStream buffer = new();
            using (XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(buffer, leaveOpen: true, new XlsbWriterOptions()))
            {
                XlsbSheetWriter sheet = workbook.AddSheet("s");
                for (int r = 0; r < 12000; r++)
                {
                    using XlsbRowWriter row = sheet.StartRow();
                    row.Write(Text(r));
                    row.Write(r);
                }
                sheet.End();
                sheet.Dispose();
                workbook.End();
            }
            return buffer.ToArray();
        }

        private static byte[] BuildXls()
        {
            using MemoryStream buffer = new();
            using (XlsWorkbookWriter workbook = XlsWorkbookWriter.Create(buffer, leaveOpen: true))
            {
                XlsSheetWriter sheet = workbook.AddSheet("s");
                for (int r = 0; r < 8000; r++)
                {
                    using XlsRowWriter row = sheet.StartRow();
                    row.Write(Text(r));
                    row.Write(r);
                }
                sheet.End();
                sheet.Dispose();
                workbook.End();
            }
            return buffer.ToArray();
        }

        [Fact]
        public async Task Large_Xlsx_Reads_The_Same_Through_A_NonSeekable_Stream()
        {
            byte[] bytes = BuildXlsx();
            Assert.True(bytes.Length > StreamSpooler.FirstChunk);
            using IExcelWorkbook expected = Excel.FromXlsx(bytes);
            using IExcelWorkbook sync = Excel.FromXlsx(new NonSeekableStream(bytes), leaveOpen: false);
            using IExcelWorkbook async = await Excel.FromXlsxAsync(new NonSeekableStream(bytes), leaveOpen: false, ct: TestContext.Current.CancellationToken);
            List<string> rows = Dump(expected);
            Assert.Equal(rows, Dump(sync));
            Assert.Equal(rows, Dump(async));
        }

        [Fact]
        public async Task Large_Xlsb_Reads_The_Same_Through_A_NonSeekable_Stream()
        {
            byte[] bytes = BuildXlsb();
            Assert.True(bytes.Length > StreamSpooler.FirstChunk);
            using IExcelWorkbook expected = Excel.FromXlsb(bytes);
            using IExcelWorkbook sync = Excel.FromXlsb(new NonSeekableStream(bytes), leaveOpen: false);
            using IExcelWorkbook async = await Excel.FromXlsbAsync(new NonSeekableStream(bytes), leaveOpen: false, ct: TestContext.Current.CancellationToken);
            List<string> rows = Dump(expected);
            Assert.Equal(rows, Dump(sync));
            Assert.Equal(rows, Dump(async));
        }

        [Fact]
        public async Task Large_Xls_Reads_The_Same_Through_A_NonSeekable_Stream()
        {
            byte[] bytes = BuildXls();
            Assert.True(bytes.Length > StreamSpooler.FirstChunk);
            using IExcelWorkbook expected = Excel.FromXls(bytes);
            using IExcelWorkbook sync = Excel.FromXls(new NonSeekableStream(bytes), leaveOpen: false);
            using IExcelWorkbook async = await Excel.FromXlsAsync(new NonSeekableStream(bytes), leaveOpen: false, ct: TestContext.Current.CancellationToken);
            List<string> rows = Dump(expected);
            Assert.Equal(rows, Dump(sync));
            Assert.Equal(rows, Dump(async));
        }
    }
}
