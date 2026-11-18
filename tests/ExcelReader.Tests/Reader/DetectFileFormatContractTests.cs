using System.IO.Compression;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests.Reader
{
    public class DetectFileFormatContractTests
    {
        private static readonly byte[] ZipSignatureOnly = [0x50, 0x4B, 0x03, 0x04];

        private static byte[] WordLikeZip()
        {
            using MemoryStream buffer = new();
            using (ZipArchive zip = new(buffer, ZipArchiveMode.Create, leaveOpen: true))
            {
                using StreamWriter writer = new(zip.CreateEntry("word/document.xml").Open());
                writer.Write("<w:document/>");
            }
            return buffer.ToArray();
        }

        private static byte[] EncryptedZip()
        {
            byte[] bytes = WordLikeZip();
            const int LocalFlagsOffset = 6;
            const int CentralFlagsOffset = 8;
            int central = bytes.AsSpan().IndexOf("PK\u0001\u0002"u8);
            Assert.True(central > 0);
            bytes[LocalFlagsOffset] |= 1;
            bytes[central + CentralFlagsOffset] |= 1;
            return bytes;
        }

        private static byte[] ZipWithMoreEntriesThanTheDefaultLimit()
        {
            using MemoryStream buffer = new();
            using (ZipArchive zip = new(buffer, ZipArchiveMode.Create, leaveOpen: true))
            {
                for (int i = 0; i <= ExcelReaderOptions.Default.MaxZipEntries; i++)
                {
                    zip.CreateEntry(i.ToString(System.Globalization.CultureInfo.InvariantCulture), CompressionLevel.NoCompression);
                }
            }
            return buffer.ToArray();
        }

        public static TheoryData<byte[]> NotAWorkbook =>
        [
            ZipSignatureOnly,
            WordLikeZip(),
            new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41, 0x41 },
        ];

        [Theory]
        [MemberData(nameof(NotAWorkbook))]
        public async Task ReturnsUnknownForAZipThatIsNotAWorkbook(byte[] bytes)
        {
            using MemoryStream stream = new(bytes);
            await using MemoryStream asyncStream = new(bytes);

            Assert.Equal(ExcelFileFormat.Unknown, Excel.DetectFileFormat(stream));
            Assert.Equal(0, stream.Position);
            Assert.Equal(ExcelFileFormat.Unknown, await Excel.DetectFileFormatAsync(asyncStream, TestContext.Current.CancellationToken));
            Assert.Equal(0, asyncStream.Position);
            Assert.Equal(ExcelFileFormat.Unknown, Excel.DetectFileFormat(bytes));
        }

        [Fact]
        public Task ReturnsUnknownForAnEncryptedZip()
        {
            return AssertUnknownAsync(EncryptedZip());
        }

        [Fact]
        public Task ReturnsUnknownForAZipPastTheEntryLimit()
        {
            return AssertUnknownAsync(ZipWithMoreEntriesThanTheDefaultLimit());
        }

        private static async Task AssertUnknownAsync(byte[] bytes)
        {
            using MemoryStream stream = new(bytes);
            await using MemoryStream asyncStream = new(bytes);

            Assert.Equal(ExcelFileFormat.Unknown, Excel.DetectFileFormat(stream));
            Assert.Equal(0, stream.Position);
            Assert.Equal(ExcelFileFormat.Unknown, await Excel.DetectFileFormatAsync(asyncStream, TestContext.Current.CancellationToken));
            Assert.Equal(0, asyncStream.Position);
            Assert.Equal(ExcelFileFormat.Unknown, Excel.DetectFileFormat(bytes));
        }

        [Fact]
        public void OpeningAZipThatIsNotAWorkbookIsRejected()
        {
            Assert.Throws<InvalidDataException>(() => Excel.Open(WordLikeZip()));
            Assert.Throws<InvalidDataException>(() => Excel.Open(new MemoryStream(WordLikeZip()), leaveOpen: false));
        }
    }
}
