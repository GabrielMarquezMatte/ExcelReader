using ExcelReader.Core.Crypto;

namespace ExcelReader.Tests
{
    // Known-answer vectors published by msoffcrypto-tool (ECMA376Standard.makekey_from_password and
    // ECMA376Standard.verifykey), an independent implementation of the same scheme. They are the
    // oracle for this derivation: the two vectors share a key, so deriving from the password and
    // then verifying with the derived key proves both halves against values this codebase did not
    // produce.
    public sealed class StandardKeyDerivationTests
    {
        private const string VectorPassword = "Password1234_";

        private static readonly byte[] VectorSalt =
        [
            0xE8, 0x82, 0x66, 0x49, 0x0C, 0x5B, 0xD1, 0xEE,
            0xBD, 0x2B, 0x43, 0x94, 0xE3, 0xF8, 0x30, 0xEF,
        ];

        private static readonly byte[] VectorKey =
        [
            0x40, 0xB1, 0x3A, 0x71, 0xF9, 0x0B, 0x96, 0x6E,
            0x37, 0x54, 0x08, 0xF2, 0xD1, 0x81, 0xA1, 0xAA,
        ];

        private static readonly byte[] VectorEncryptedVerifier =
        [
            0x51, 0x6F, 0x73, 0x2E, 0x96, 0x6F, 0xAC, 0x17,
            0xB1, 0xC5, 0xD7, 0xD8, 0xCC, 0x36, 0xC9, 0x28,
        ];

        private static readonly byte[] VectorEncryptedVerifierHash =
        [
            0x2B, 0x61, 0x68, 0xDA, 0xBE, 0x29, 0x11, 0xAD,
            0x2B, 0xD3, 0x7C, 0x17, 0x46, 0x74, 0x5C, 0x14,
            0xD3, 0xCF, 0x1B, 0xB1, 0x40, 0xA4, 0x8F, 0x4E,
            0x6F, 0x3D, 0x23, 0x88, 0x08, 0x72, 0xB1, 0x6A,
        ];

        [Fact]
        public void Should_Match_Published_Vector_When_Deriving_An_Aes128_Key()
        {
            byte[] key = StandardKeyDerivation.DeriveKey(VectorPassword, VectorSalt, keyBits: 128);

            Assert.Equal(VectorKey, key);
        }

        [Fact]
        public void Should_Accept_When_Verifier_Matches_The_Derived_Key()
        {
            byte[] key = StandardKeyDerivation.DeriveKey(VectorPassword, VectorSalt, keyBits: 128);

            Assert.True(StandardKeyDerivation.VerifyPassword(
                key, VectorEncryptedVerifier, VectorEncryptedVerifierHash));
        }

        [Fact]
        public void Should_Reject_When_Verifier_Hash_Is_Tampered()
        {
            byte[] key = StandardKeyDerivation.DeriveKey(VectorPassword, VectorSalt, keyBits: 128);
            byte[] tampered = [.. VectorEncryptedVerifierHash];
            tampered[0] ^= 0xFF;

            Assert.False(StandardKeyDerivation.VerifyPassword(key, VectorEncryptedVerifier, tampered));
        }

        [Fact]
        public void Should_Reject_When_The_Password_Is_Wrong()
        {
            byte[] key = StandardKeyDerivation.DeriveKey("not-the-password", VectorSalt, keyBits: 128);

            Assert.False(StandardKeyDerivation.VerifyPassword(
                key, VectorEncryptedVerifier, VectorEncryptedVerifierHash));
        }

        [Theory]
        [InlineData(128, 16)]
        [InlineData(192, 24)]
        [InlineData(256, 32)]
        public void Should_Return_Requested_Length_When_KeyBits_Vary(int keyBits, int expectedBytes)
        {
            byte[] key = StandardKeyDerivation.DeriveKey(VectorPassword, VectorSalt, keyBits);

            Assert.Equal(expectedBytes, key.Length);
        }

        [Fact]
        public void Should_Return_False_When_EncryptedVerifierHash_Exceeds_The_Cipher_Block_Buffer()
        {
            byte[] key = StandardKeyDerivation.DeriveKey(VectorPassword, VectorSalt, keyBits: 128);
            byte[] oversizedHash = new byte[96];

            Assert.False(StandardKeyDerivation.VerifyPassword(key, VectorEncryptedVerifier, oversizedHash));
        }
    }
}
