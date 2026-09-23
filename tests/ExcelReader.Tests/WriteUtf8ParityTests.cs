using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    public sealed class WriteUtf8ParityTests
    {
        public static TheoryData<byte[]> Values =>
        [
            "plain"u8.ToArray(),
            "café 日本 😀"u8.ToArray(),
            "a&b<c>d\"e'f"u8.ToArray(),
            "_x0041_ stays literal"u8.ToArray(),
            "under_score"u8.ToArray(),
            " leading"u8.ToArray(),
            "trailing\t"u8.ToArray(),
            " nbsp edge"u8.ToArray(),
            "ctrl\u0001char"u8.ToArray(),
            "line\nbreak,comma"u8.ToArray(),
            "quote\"inside"u8.ToArray(),
            Array.Empty<byte>(),
            new byte[] { 0x61, 0xFF, 0x62 },
            new byte[] { 0xC3 },
        ];

        [Theory]
        [MemberData(nameof(Values))]
        public void XlsxInlineWriteUtf8MatchesWriteString(byte[] utf8)
        {
            Assert.Equal(XlsxSheetXml(utf8, useUtf8: false, sharedStrings: false), XlsxSheetXml(utf8, useUtf8: true, sharedStrings: false));
        }

        [Theory]
        [MemberData(nameof(Values))]
        public void XlsxSharedStringsWriteUtf8MatchesWriteString(byte[] utf8)
        {
            Assert.Equal(XlsxSheetXml(utf8, useUtf8: false, sharedStrings: true), XlsxSheetXml(utf8, useUtf8: true, sharedStrings: true));
        }

        [Theory]
        [MemberData(nameof(Values))]
        public void CsvWriteUtf8MatchesWriteString(byte[] utf8)
        {
            Assert.Equal(CsvBytes(utf8, useUtf8: false), CsvBytes(utf8, useUtf8: true));
        }

        [Fact]
        public void XlsxWriteUtf8RejectsTextOverTheCellLimitLikeWriteString()
        {
            byte[] utf8 = Encoding.UTF8.GetBytes(new string('x', 32_768));
            Assert.Throws<ArgumentException>(() => XlsxSheetXml(utf8, useUtf8: true, sharedStrings: false));
        }

        private static string XlsxSheetXml(byte[] utf8, bool useUtf8, bool sharedStrings)
        {
            using MemoryStream stream = new();
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(stream, leaveOpen: true, options: new XlsxWriterOptions { UseSharedStrings = sharedStrings }))
            {
                using XlsxSheetWriter sheet = workbook.AddSheet("S");
                sheet.Start();
                using (XlsxRowWriter row = sheet.StartRow())
                {
                    Write(row, utf8, useUtf8);
                    row.Write(1);
                }
                sheet.End();
                workbook.End();
            }
            stream.Position = 0;
            using ZipArchive zip = new(stream, ZipArchiveMode.Read);
            StringBuilder parts = new();
            foreach (string name in (string[])["xl/worksheets/sheet1.xml", "xl/sharedStrings.xml"])
            {
                ZipArchiveEntry? entry = zip.GetEntry(name);
                if (entry is null)
                {
                    continue;
                }
                using StreamReader reader = new(entry.Open());
                parts.Append(reader.ReadToEnd());
            }
            return parts.ToString();
        }

        private static byte[] CsvBytes(byte[] utf8, bool useUtf8)
        {
            using MemoryStream stream = new();
            using (CsvWorkbookWriter workbook = CsvWorkbookWriter.Create(stream, leaveOpen: true))
            {
                using CsvSheetWriter sheet = workbook.AddSheet("S");
                sheet.Start();
                using (CsvRowWriter row = sheet.StartRow())
                {
                    Write(row, utf8, useUtf8);
                    row.Write(1);
                }
                sheet.End();
                workbook.End();
            }
            return stream.ToArray();
        }

        private static void Write(IRowWriter row, byte[] utf8, bool useUtf8)
        {
            if (useUtf8)
            {
                row.WriteUtf8(utf8);
            }
            else
            {
                row.Write(Encoding.UTF8.GetString(utf8));
            }
        }
    }
}
