using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    // EncryptionInfo is attacker-controlled and parsed before any password check, so every negative
    // case here forges a specific malformed field and asserts the exact rejection. The fuzz oracle
    // treats an index/range exception from this path as a defect, so "it threw something" is not
    // good enough.
    public sealed class StandardDescriptorTests
    {
        private const int AlgIdAes128 = 0x0000660E;
        private const int AlgIdAes256 = 0x00006610;
        private const int AlgIdRc4 = 0x00006801;

        private static EncryptionDescriptor Parse(byte[] info)
        {
            return EncryptionDescriptor.Parse(info, ExcelReaderOptions.Default);
        }

        [Theory]
        [InlineData(AlgIdAes128, 128)]
        [InlineData(0x0000660F, 192)]
        [InlineData(AlgIdAes256, 256)]
        public void Should_Parse_When_Header_Is_Well_Formed(int algId, int keySize)
        {
            StandardDescriptor descriptor = Assert.IsType<StandardDescriptor>(
                Parse(StandardInfoBuilder.Build(algId: algId, keySize: keySize)));

            Assert.Equal(algId, descriptor.AlgId);
            Assert.Equal(keySize, descriptor.KeyBits);
            Assert.Equal(16, descriptor.Salt.Length);
            Assert.Equal(16, descriptor.EncryptedVerifier.Length);
            Assert.Equal(32, descriptor.EncryptedVerifierHash.Length);
        }

        [Fact]
        public void Should_Infer_KeyBits_From_AlgId_When_KeySize_Is_Zero()
        {
            StandardDescriptor descriptor = Assert.IsType<StandardDescriptor>(
                Parse(StandardInfoBuilder.Build(algId: AlgIdAes128, keySize: 0)));

            Assert.Equal(128, descriptor.KeyBits);
        }

        [Fact]
        public void Should_Reject_As_Rc4_When_AlgId_Is_Rc4()
        {
            ExcelEncryptionException ex = Assert.Throws<ExcelEncryptionException>(
                () => Parse(StandardInfoBuilder.Build(algId: AlgIdRc4, keySize: 128)));

            Assert.Equal(ExcelEncryptionReason.UnsupportedScheme, ex.Reason);
            Assert.Contains("RC4", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Should_Reject_When_AlgId_Is_Unrecognized()
        {
            ExcelEncryptionException ex = Assert.Throws<ExcelEncryptionException>(
                () => Parse(StandardInfoBuilder.Build(algId: 0x00001234)));

            Assert.Equal(ExcelEncryptionReason.UnsupportedScheme, ex.Reason);
        }

        [Fact]
        public void Should_Reject_When_HashAlgorithm_Is_Not_Sha1()
        {
            ExcelEncryptionException ex = Assert.Throws<ExcelEncryptionException>(
                () => Parse(StandardInfoBuilder.Build(algIdHash: 0x0000800C)));

            Assert.Equal(ExcelEncryptionReason.UnsupportedScheme, ex.Reason);
        }

        [Fact]
        public void Should_Accept_When_HashAlgorithm_Is_The_Scheme_Default()
        {
            Assert.IsType<StandardDescriptor>(Parse(StandardInfoBuilder.Build(algIdHash: 0)));
        }

        [Fact]
        public void Should_Reject_When_KeySize_Disagrees_With_AlgId()
        {
            ExcelEncryptionException ex = Assert.Throws<ExcelEncryptionException>(
                () => Parse(StandardInfoBuilder.Build(algId: AlgIdAes256, keySize: 128)));

            Assert.Equal(ExcelEncryptionReason.UnsupportedScheme, ex.Reason);
        }

        [Theory]
        [InlineData(4)]     // shorter than the eight fixed header fields
        [InlineData(31)]
        public void Should_Reject_When_HeaderSize_Is_Too_Small(int headerSize)
        {
            Assert.Throws<InvalidDataException>(
                () => Parse(StandardInfoBuilder.Build(headerSizeOverride: headerSize)));
        }

        [Fact]
        public void Should_Reject_When_HeaderSize_Exceeds_The_Stream()
        {
            Assert.Throws<InvalidDataException>(
                () => Parse(StandardInfoBuilder.Build(headerSizeOverride: 100_000)));
        }

        [Fact]
        public void Should_Reject_When_SaltSize_Is_Not_Sixteen()
        {
            Assert.Throws<InvalidDataException>(() => Parse(StandardInfoBuilder.Build(saltSize: 8)));
        }

        [Theory]
        [InlineData(20)]
        [InlineData(32)]
        [InlineData(0)]
        public void Should_Parse_When_VerifierHashSize_Is_Any_Declared_Value(int verifierHashSize)
        {
            Assert.IsType<StandardDescriptor>(
                Parse(StandardInfoBuilder.Build(verifierHashSize: verifierHashSize)));
        }

        [Fact]
        public void Should_Reject_When_Verifier_Region_Is_Truncated()
        {
            byte[] info = StandardInfoBuilder.Build();
            Assert.Throws<InvalidDataException>(() => Parse(info[..(info.Length - 8)]));
        }

        [Fact]
        public void Should_Reject_When_Stream_Ends_Before_The_Header_Size_Field()
        {
            Assert.Throws<InvalidDataException>(() => Parse(StandardInfoBuilder.Build()[..10]));
        }
    }
}
