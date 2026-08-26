using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class CfbContainerTests
    {
        private static CfbContainer Open(string fixture)
        {
            var fs = File.OpenRead(EncryptedFixtures.Path_(fixture));
            return CfbContainer.Parse(fs, ownsSource: true, ExcelReaderOptions.Default);
        }

        [Fact]
        public void Should_Find_Both_Encryption_Streams_When_Encrypted_Container()
        {
            using CfbContainer cfb = Open("agile-aes256-sha512.xlsx");
            Assert.True(cfb.ContainsStream("EncryptionInfo"));
            Assert.True(cfb.ContainsStream("EncryptedPackage"));
            Assert.False(cfb.ContainsStream("Workbook"));
        }

        // EncryptionInfo is a few hundred bytes and sits below the 4096-byte mini-stream cutoff,
        // so this also exercises the mini-FAT path.
        [Fact]
        public void Should_Read_EncryptionInfo_When_Below_Mini_Cutoff()
        {
            using CfbContainer cfb = Open("agile-aes256-sha512.xlsx");
            byte[] info = cfb.ReadStream("EncryptionInfo", maxBytes: 64 * 1024);
            Assert.InRange(info.Length, 64, 64 * 1024);
            Assert.Equal(4, info[0] | (info[1] << 8));   // major version 4 = agile
            Assert.Equal(4, info[2] | (info[3] << 8));   // minor version 4 = agile
        }

        [Fact]
        public void Should_Throw_When_Stream_Exceeds_Max_Bytes()
        {
            using CfbContainer cfb = Open("agile-aes256-sha512.xlsx");
            Assert.Throws<ExcelLimitExceededException>(
                () => cfb.ReadStream("EncryptionInfo", maxBytes: 8));
        }

        [Fact]
        public void Should_Throw_When_Stream_Absent()
        {
            using CfbContainer cfb = Open("agile-aes256-sha512.xlsx");
            Assert.Throws<InvalidDataException>(() => cfb.ReadStream("NoSuchStream", 1024));
        }

        // The EncryptedPackage view is what ZipArchive will eventually seek over, so seekability
        // and an accurate Length are load-bearing, not incidental.
        [Fact]
        public void Should_Expose_Seekable_View_When_Opening_EncryptedPackage()
        {
            using CfbContainer cfb = Open("agile-aes256-sha512.xlsx");
            using Stream view = cfb.OpenStreamView("EncryptedPackage");
            Assert.True(view.CanRead);
            Assert.True(view.CanSeek);
            Assert.False(view.CanWrite);
            Assert.Equal(cfb.StreamLength("EncryptedPackage"), view.Length);

            // The first 8 bytes are the little-endian plaintext size; the ciphertext follows.
            byte[] prefix = new byte[8];
            view.ReadExactly(prefix);
            long declared = BitConverter.ToInt64(prefix);
            Assert.InRange(declared, 1, view.Length);
        }

        // Reading the tail via Seek must agree with reading it sequentially - the whole point of
        // the view is that ZipArchive can jump to the central directory.
        [Fact]
        public void Should_Match_Sequential_Read_When_Seeking()
        {
            using CfbContainer cfb = Open("agile-aes256-sha512.xlsx");
            using Stream view = cfb.OpenStreamView("EncryptedPackage");
            byte[] whole = new byte[view.Length];
            view.ReadExactly(whole);

            view.Seek(-64, SeekOrigin.End);
            byte[] tail = new byte[64];
            view.ReadExactly(tail);
            Assert.Equal(whole[^64..], tail);
        }
    }
}
