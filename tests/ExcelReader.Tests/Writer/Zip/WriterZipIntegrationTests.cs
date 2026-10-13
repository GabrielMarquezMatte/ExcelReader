using System.Globalization;
using System.IO.Compression;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer.Zip
{
    public class WriterZipIntegrationTests
    {
        private const int Rows = 30_000;

        public static TheoryData<CompressionLevel, bool, bool> Cases()
        {
            var data = new TheoryData<CompressionLevel, bool, bool>();
            foreach (CompressionLevel level in Enum.GetValues<CompressionLevel>())
            {
                foreach (bool seekable in new[] { true, false })
                {
                    data.Add(level, seekable, false);
                    data.Add(level, seekable, true);
                }
            }
            return data;
        }

        private static Stream Target(MemoryStream ms, bool seekable)
        {
            if (seekable)
            {
                return ms;
            }
            return new ForwardOnlyWriteStream(ms);
        }

        private static void AssertEveryEntryCrc(byte[] package)
        {
            using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            Assert.NotEmpty(zip.Entries);
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                using Stream data = entry.Open();
                using var copy = new MemoryStream();
                data.CopyTo(copy);
                Assert.Equal(Crc32Tests.BitwiseCrc(copy.ToArray()), entry.Crc32);
            }
        }

        private static void AssertRows(IExcelWorkbook reader)
        {
            using IExcelRowEnumerator rows = reader.SheetAt(0).GetEnumerator();
            for (int i = 0; i < Rows; i++)
            {
                Assert.True(rows.MoveNext());
                Assert.Equal(i.ToString(CultureInfo.InvariantCulture), rows.Current[0].GetString());
                Assert.Equal($"text {i % 101}", rows.Current[1].GetString());
            }
            Assert.False(rows.MoveNext());
            using IExcelRowEnumerator second = reader.SheetAt(1).GetEnumerator();
            Assert.True(second.MoveNext());
            Assert.Equal("second", second.Current[0].GetString());
        }

        [Theory]
        [MemberData(nameof(Cases))]
        public async Task XlsxReadsBackFromEveryLevelAndOutput(CompressionLevel level, bool seekable, bool prefetch)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var ms = new MemoryStream();
            var options = new XlsxWriterOptions { Compression = level, PrefetchWrite = prefetch, UseSharedStrings = true };
            await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(Target(ms, seekable), leaveOpen: true, options))
            {
                XlsxSheetWriter sheet = wb.AddSheet("Data");
                for (int i = 0; i < Rows; i++)
                {
                    await using XlsxRowWriter row = await sheet.StartRowAsync(ct);
                    row.Write(i);
                    row.Write($"text {i % 101}");
                    row.Write(i * 1.5);
                }
                await sheet.EndAsync(ct);
                XlsxSheetWriter second = wb.AddSheet("Second");
                await using (XlsxRowWriter row = await second.StartRowAsync(ct))
                {
                    row.Write("second");
                }
                await second.EndAsync(ct);
                await wb.EndAsync(ct);
            }
            AssertEveryEntryCrc(ms.ToArray());
            using XlsxWorkbook reader = Excel.FromXlsx(new MemoryStream(ms.ToArray()));
            AssertRows(reader);
        }

        [Theory]
        [MemberData(nameof(Cases))]
        public void XlsbReadsBackFromEveryLevelAndOutput(CompressionLevel level, bool seekable, bool prefetch)
        {
            var ms = new MemoryStream();
            var options = new XlsbWriterOptions { Compression = level, PrefetchWrite = prefetch, UseSharedStrings = true };
            using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(Target(ms, seekable), leaveOpen: true, options))
            {
                XlsbSheetWriter sheet = wb.AddSheet("Data");
                for (int i = 0; i < Rows; i++)
                {
                    using XlsbRowWriter row = sheet.StartRow();
                    row.Write(i);
                    row.Write($"text {i % 101}");
                    row.Write(i * 1.5);
                }
                sheet.End();
                XlsbSheetWriter second = wb.AddSheet("Second");
                using (XlsbRowWriter row = second.StartRow())
                {
                    row.Write("second");
                }
                second.End();
                wb.End();
            }
            AssertEveryEntryCrc(ms.ToArray());
            using XlsbWorkbook reader = Excel.FromXlsb(new MemoryStream(ms.ToArray()));
            AssertRows(reader);
        }
    }
}
