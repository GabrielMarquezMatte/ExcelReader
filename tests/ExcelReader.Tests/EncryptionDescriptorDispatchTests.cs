using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    // The version tuple alone does not identify the cipher: [MS-OFFCRYPTO] allows 2.2/3.2/4.2 for
    // both standard AES and RC4 CryptoAPI, and EncryptionHeader.algId is what distinguishes them.
    // These cases pin that, including the two combinations the previous version-only dispatch named
    // wrongly.
    public sealed class EncryptionDescriptorDispatchTests
    {
        private static EncryptionDescriptor Parse(byte[] info)
        {
            return EncryptionDescriptor.Parse(info, ExcelReaderOptions.Default);
        }

        private static byte[] Version(int major, int minor)
        {
            byte[] info = new byte[8];
            info[0] = (byte)major;
            info[2] = (byte)minor;
            return info;
        }

        [Theory]
        [InlineData(2, 2)]
        [InlineData(3, 2)]
        [InlineData(4, 2)]
        public void Should_Route_To_Standard_When_Version_Is_Binary_And_AlgId_Is_Aes(int major, int minor)
        {
            byte[] info = StandardInfoBuilder.Build(major, minor, algId: 0x00006610, keySize: 256);

            Assert.IsType<StandardDescriptor>(Parse(info));
        }

        [Fact]
        public void Should_Name_Rc4_When_Version_Is_Four_Two_And_AlgId_Is_Rc4()
        {
            byte[] info = StandardInfoBuilder.Build(4, 2, algId: 0x00006801, keySize: 128);

            ExcelEncryptionException ex = Assert.Throws<ExcelEncryptionException>(() => Parse(info));

            Assert.Equal(ExcelEncryptionReason.UnsupportedScheme, ex.Reason);
            Assert.Contains("RC4", ex.Message, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(3, 1)]
        [InlineData(4, 1)]
        public void Should_Name_Rc4_When_Version_Is_A_Rc4_CryptoApi_Tuple(int major, int minor)
        {
            ExcelEncryptionException ex = Assert.Throws<ExcelEncryptionException>(
                () => Parse(Version(major, minor)));

            Assert.Equal(ExcelEncryptionReason.UnsupportedScheme, ex.Reason);
            Assert.Contains("RC4", ex.Message, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(3, 3)]
        [InlineData(4, 3)]
        public void Should_Name_Extensible_When_Version_Is_An_Extensible_Tuple(int major, int minor)
        {
            ExcelEncryptionException ex = Assert.Throws<ExcelEncryptionException>(
                () => Parse(Version(major, minor)));

            Assert.Equal(ExcelEncryptionReason.UnsupportedScheme, ex.Reason);
            Assert.Contains("extensible", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        // Undefined minors under major 2 used to be swept into the RC4 message by a (2, _)
        // catch-all. They are not RC4; they are not any defined scheme.
        [Theory]
        [InlineData(2, 0)]
        [InlineData(2, 1)]
        [InlineData(2, 3)]
        [InlineData(5, 5)]
        public void Should_Name_The_Version_When_Tuple_Is_Undefined(int major, int minor)
        {
            ExcelEncryptionException ex = Assert.Throws<ExcelEncryptionException>(
                () => Parse(Version(major, minor)));

            Assert.Equal(ExcelEncryptionReason.UnsupportedScheme, ex.Reason);
            Assert.Contains("Unrecognized", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Should_Reject_When_Stream_Is_Shorter_Than_The_Version_Tuple()
        {
            Assert.Throws<InvalidDataException>(() => Parse([0x04, 0x00]));
        }
    }
}
