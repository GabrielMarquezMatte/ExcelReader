using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Reader.Zip
{
    /// <summary>A package whose entry data no longer matches its CRC-32 is refused, unless the check is turned off.</summary>
    public sealed class ZipChecksumTests
    {
        private const int Rows = 20_000;

        private static readonly ExcelReaderOptions Unchecked = new() { VerifyZipChecksums = false };

        public static TheoryData<bool, bool, bool> Cases
        {
            get
            {
                var data = new TheoryData<bool, bool, bool>();
                foreach (bool xlsb in new[] { false, true })
                {
                    foreach (bool stream in new[] { false, true })
                    {
                        data.Add(xlsb, stream, false);
                        data.Add(xlsb, stream, true);
                    }
                }
                return data;
            }
        }

        [Theory]
        [MemberData(nameof(Cases))]
        public void ADamagedSheetIsRefused(bool xlsb, bool fromStream, bool prefetch)
        {
            byte[] package = Damage(Write(xlsb), xlsb, xlsb ? "xl/worksheets/sheet1.bin" : "xl/worksheets/sheet1.xml");
            var options = new ExcelReaderOptions { PrefetchDecompression = prefetch };

            InvalidDataException error = Assert.Throws<InvalidDataException>(() => Sum(package, fromStream, options));
            Assert.Contains("CRC-32", error.Message, StringComparison.Ordinal);
        }

        [Theory]
        [MemberData(nameof(Cases))]
        public void TheCheckCanBeTurnedOff(bool xlsb, bool fromStream, bool prefetch)
        {
            byte[] package = Damage(Write(xlsb), xlsb, xlsb ? "xl/worksheets/sheet1.bin" : "xl/worksheets/sheet1.xml");

            Sum(package, fromStream, Unchecked with { PrefetchDecompression = prefetch });
        }

        [Theory]
        [MemberData(nameof(Cases))]
        public void AnIntactPackageStillReads(bool xlsb, bool fromStream, bool prefetch)
        {
            long sum = Sum(Write(xlsb), fromStream, new ExcelReaderOptions { PrefetchDecompression = prefetch });

            Assert.Equal((long)Rows * (Rows + 1) / 2, sum);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void ADamagedSharedStringTableIsRefused(bool xlsb)
        {
            byte[] package = Damage(Write(xlsb, sharedStrings: true), xlsb, xlsb ? "xl/sharedStrings.bin" : "xl/sharedStrings.xml");

            Assert.Throws<InvalidDataException>(() => Sum(package, fromStream: false, ExcelReaderOptions.Default));
            Assert.Throws<InvalidDataException>(() => Sum(package, fromStream: true, ExcelReaderOptions.Default));
        }

        private static long Sum(byte[] package, bool fromStream, ExcelReaderOptions options)
        {
            using IExcelWorkbook workbook = fromStream
                ? Excel.Open(new MemoryStream(package, writable: false), leaveOpen: false, options)
                : Excel.Open(package, options);
            long sum = 0;
            using IExcelRowEnumerator rows = workbook.FirstSheet.GetEnumerator();
            while (rows.MoveNext())
            {
                rows.Current[0].GetString();
                Assert.True(rows.Current[1].TryGetDouble(out double number));
                sum += (long)number;
            }
            return sum;
        }

        private static byte[] Write(bool xlsb, bool sharedStrings = false)
        {
            using MemoryStream ms = new();
            if (xlsb)
            {
                using XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true, new XlsbWriterOptions { Compression = CompressionLevel.NoCompression, UseSharedStrings = sharedStrings });
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                for (int i = 1; i <= Rows; i++)
                {
                    using XlsbRowWriter row = sheet.StartRow();
                    row.Write("text" + (i % 50));
                    row.Write(i);
                }
                sheet.End();
                wb.End();
            }
            else
            {
                using XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true, new XlsxWriterOptions { Compression = CompressionLevel.NoCompression, UseSharedStrings = sharedStrings });
                XlsxSheetWriter sheet = wb.AddSheet("S1");
                for (int i = 1; i <= Rows; i++)
                {
                    using XlsxRowWriter row = sheet.StartRow();
                    row.Write("text" + (i % 50));
                    row.Write(i);
                }
                sheet.End();
                wb.End();
            }
            return ms.ToArray();
        }

        // The package is stored, not deflated, so one changed letter of the cell text "text" stays valid
        // XML or BIFF12 and only the checksum can tell.
        private static byte[] Damage(byte[] package, bool xlsb, string entryName)
        {
            byte[] content;
            using (var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read))
            {
                using Stream data = zip.GetEntry(entryName)!.Open();
                using var copy = new MemoryStream();
                data.CopyTo(copy);
                content = copy.ToArray();
            }
            // BIFF12 stores text as UTF-16.
            byte[] needle = xlsb ? Encoding.Unicode.GetBytes("text") : Encoding.ASCII.GetBytes("text");
            int inEntry = Find(content, needle, content.Length / 2);
            int inPackage = Find(package, content.AsSpan(inEntry - 16, 64), 0) + 16;
            byte[] damaged = (byte[])package.Clone();
            damaged[inPackage] = (byte)'T';
            return damaged;
        }

        private static int Find(ReadOnlySpan<byte> haystack, ReadOnlySpan<byte> needle, int from)
        {
            int at = haystack[from..].IndexOf(needle);
            if (at < 0 && from > 0)
            {
                at = haystack.IndexOf(needle) - from;
            }
            Assert.True(from + at >= 0, "needle not found");
            return from + at;
        }
    }
}
