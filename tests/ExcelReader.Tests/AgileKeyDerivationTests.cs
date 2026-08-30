using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class AgileKeyDerivationTests
    {
        private static AgileDescriptor Descriptor(string fixture)
        {
            var fs = File.OpenRead(EncryptedFixtures.Path_(fixture));
            using CfbContainer cfb = CfbContainer.Parse(fs, ownsSource: true, ExcelReaderOptions.Default);
            byte[] info = cfb.ReadStream("EncryptionInfo", 64 * 1024);
            return (AgileDescriptor)EncryptionDescriptor.Parse(info, ExcelReaderOptions.Default);
        }

        public static TheoryData<string> AgileFixtures()
        {
            return new()
        {
            "agile-aes256-sha512.xlsx",
            "agile-aes256-sha512.xlsb",
        };
        }

        // Only AES-256/SHA-512 has a real fixture in this pass (see "Execution Scope Note"); Task 3's
        // Should_Parse_Agile_When_KeyBits_And_Hash_Vary pins the AES-128/SHA-1 field layout separately,
        // without a derivation oracle to check it against.
        [Theory]
        [MemberData(nameof(AgileFixtures))]
        public void Should_Derive_Key_When_Password_Correct(string fixture)
        {
            AgileDescriptor d = Descriptor(fixture);
            byte[] key = AgileKeyDerivation.DeriveIntermediateKey(d, EncryptedFixtures.Password);
            Assert.Equal(d.KeyData.KeyBits / 8, key.Length);
        }

        [Theory]
        [MemberData(nameof(AgileFixtures))]
        public void Should_Report_PasswordIncorrect_When_Password_Wrong(string fixture)
        {
            AgileDescriptor d = Descriptor(fixture);
            var ex = Assert.Throws<ExcelEncryptionException>(
                () => AgileKeyDerivation.DeriveIntermediateKey(d, "not-the-password"));
            Assert.Equal(ExcelEncryptionReason.PasswordIncorrect, ex.Reason);
        }

        // EncryptionInfo is untrusted, attacker-controlled input parsed before any password check.
        // A ciphertext length that isn't a multiple of the AES block size must be reported as
        // malformed input (InvalidDataException), not let AES's own CryptographicException escape —
        // callers of DeriveIntermediateKey only expect ExcelEncryptionException/InvalidDataException,
        // and a raw CryptographicException would also fail FuzzMutation.AcceptableExceptionTypes.
        [Fact]
        public void Should_Throw_InvalidData_When_Ciphertext_Length_Is_Misaligned()
        {
            AgileDescriptor d = Descriptor("agile-aes256-sha512.xlsx");
            byte[] truncated = d.PasswordEncryptor.EncryptedVerifierHashInput[..^1];
            CryptoParameters malformed = d.PasswordEncryptor with { EncryptedVerifierHashInput = truncated };
            AgileDescriptor tampered = d with { PasswordEncryptor = malformed };

            Assert.Throws<InvalidDataException>(
                () => AgileKeyDerivation.DeriveIntermediateKey(tampered, EncryptedFixtures.Password));
        }

        // encryptedKeyValue is untrusted, pre-authentication input parsed before any password check.
        // The spec defines its plaintext to be exactly KeyData.keyBits/8 bytes (32 here); a producer
        // that wraps fewer bytes than that must be reported as malformed input (InvalidDataException),
        // not let `raw[..keyLen]` throw a raw ArgumentOutOfRangeException. Truncating to one AES block
        // (16 of the expected 32 bytes) keeps the ciphertext itself block-aligned - so it still
        // decrypts cleanly under CBC - while decrypting to fewer bytes than KeyBits declares.
        [Fact]
        public void Should_Throw_InvalidData_When_EncryptedKeyValue_Decrypts_Shorter_Than_KeyBits()
        {
            AgileDescriptor d = Descriptor("agile-aes256-sha512.xlsx");
            byte[] shortKeyValue = d.PasswordEncryptor.EncryptedKeyValue[..16];
            CryptoParameters malformed = d.PasswordEncryptor with { EncryptedKeyValue = shortKeyValue };
            AgileDescriptor tampered = d with { PasswordEncryptor = malformed };

            Assert.Throws<InvalidDataException>(
                () => AgileKeyDerivation.DeriveIntermediateKey(tampered, EncryptedFixtures.Password));
        }

        // Each 4096-byte segment gets its own IV derived from its index; if these collided, the
        // multi-segment fixture would decrypt to garbage past the first segment.
        [Fact]
        public void Should_Produce_Distinct_Ivs_Per_Segment()
        {
            AgileDescriptor d = Descriptor("agile-aes256-sha512.xlsx");
            byte[] iv0 = AgileKeyDerivation.SegmentIv(d, 0);
            byte[] iv1 = AgileKeyDerivation.SegmentIv(d, 1);
            Assert.Equal(d.KeyData.BlockSize, iv0.Length);
            Assert.NotEqual(iv0, iv1);
            Assert.Equal(iv0, AgileKeyDerivation.SegmentIv(d, 0));
        }

        [Fact]
        public void Should_Unwrap_Hmac_Key_When_Descriptor_Has_DataIntegrity()
        {
            AgileDescriptor d = Descriptor("agile-aes256-sha512.xlsx");
            byte[] key = AgileKeyDerivation.DeriveIntermediateKey(d, EncryptedFixtures.Password);
            (byte[] hmacKey, byte[] hmacValue) = AgileKeyDerivation.UnwrapHmac(d, key);
            Assert.Equal(AgileKeyDerivation.HashLength(d.KeyData.Hash), hmacValue.Length);
            Assert.NotEmpty(hmacKey);
        }

        // A crafted descriptor can ask for billions of iterations; that must be a bounded rejection,
        // not an hours-long stall.
        [Fact]
        public void Should_Throw_When_SpinCount_Exceeds_Limit()
        {
            var fs = File.OpenRead(EncryptedFixtures.Path_("agile-aes256-sha512.xlsx"));
            using CfbContainer cfb = CfbContainer.Parse(fs, ownsSource: true, ExcelReaderOptions.Default);
            byte[] info = cfb.ReadStream("EncryptionInfo", 64 * 1024);
            var tight = ExcelReaderOptions.Default with { MaxPasswordSpinCount = 1 };
            Assert.Throws<ExcelLimitExceededException>(() => EncryptionDescriptor.Parse(info, tight));
        }
    }
}
