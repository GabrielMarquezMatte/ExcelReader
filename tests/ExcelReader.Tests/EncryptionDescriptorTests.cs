using System.Text;
using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class EncryptionDescriptorTests
    {
        private static byte[] Info(string fixture)
        {
            var fs = File.OpenRead(EncryptedFixtures.Path_(fixture));
            using CfbContainer cfb = CfbContainer.Parse(fs, ownsSource: true, ExcelReaderOptions.Default);
            return cfb.ReadStream("EncryptionInfo", maxBytes: 64 * 1024);
        }

        [Fact]
        public void Should_Parse_Agile_When_Version_4_4()
        {
            var d = Assert.IsType<AgileDescriptor>(
                EncryptionDescriptor.Parse(Info("agile-aes256-sha512.xlsx"), ExcelReaderOptions.Default));
            Assert.Equal(256, d.KeyData.KeyBits);
            Assert.Equal(HashKind.Sha512, d.KeyData.Hash);
            Assert.Equal(16, d.KeyData.BlockSize);
            Assert.NotEmpty(d.KeyData.SaltValue);
            Assert.NotEmpty(d.PasswordEncryptor.EncryptedKeyValue);
            Assert.True(d.PasswordEncryptor.SpinCount > 0);
        }

        // AES-128/SHA-1 agile parsing has no corresponding real fixture in this pass (see
        // "Execution Scope Note"), so this pins the field layout against hand-built bytes instead
        // of asserting decryption correctness — Task 5's derivation tests are what need a real
        // oracle, and they only run against agile-aes256-sha512.*.
        [Fact]
        public void Should_Parse_Agile_When_KeyBits_And_Hash_Vary()
        {
            byte[] info = Info("agile-aes256-sha512.xlsx");
            string xml = Encoding.UTF8.GetString(info.AsSpan(8))
                .Replace("keyBits=\"256\"", "keyBits=\"128\"", StringComparison.Ordinal)
                .Replace("hashAlgorithm=\"SHA512\"", "hashAlgorithm=\"SHA1\"", StringComparison.Ordinal)
                .Replace("hashSize=\"64\"", "hashSize=\"20\"", StringComparison.Ordinal);
            byte[] patched = [.. info.AsSpan(0, 8), .. Encoding.UTF8.GetBytes(xml)];
            var d = Assert.IsType<AgileDescriptor>(EncryptionDescriptor.Parse(patched, ExcelReaderOptions.Default));
            Assert.Equal(128, d.KeyData.KeyBits);
            Assert.Equal(HashKind.Sha1, d.KeyData.Hash);
        }

        // Standard encryption (3.2/4.2) parses into a StandardDescriptor (see
        // StandardDescriptorTests/EncryptionDescriptorDispatchTests for the full coverage) and, via
        // DecryptedPackageStream.Create/StandardPackageCipher, decrypts fully. A well-formed binary
        // header succeeds at Parse() here; the decrypt path itself is covered elsewhere.
        [Fact]
        public void Should_Parse_When_Standard_Encryption_Is_Well_Formed()
        {
            foreach ((int major, int minor) in new[] { (3, 2), (4, 2) })
            {
                byte[] info = StandardInfoBuilder.Build(major, minor);
                Assert.IsType<StandardDescriptor>(EncryptionDescriptor.Parse(info, ExcelReaderOptions.Default));
            }
        }

        // Version 2.2 is a binary-header version shared with standard AES; EncryptionHeader.algId,
        // not the version tuple, is what names it RC4 CryptoAPI - explicitly out of scope, and no
        // password will ever help.
        [Fact]
        public void Should_Report_UnsupportedScheme_When_Rc4_CryptoApi()
        {
            byte[] rc4 = StandardInfoBuilder.Build(major: 2, minor: 2, algId: 0x00006801, keySize: 128);
            var ex = Assert.Throws<ExcelEncryptionException>(
                () => EncryptionDescriptor.Parse(rc4, ExcelReaderOptions.Default));
            Assert.Equal(ExcelEncryptionReason.UnsupportedScheme, ex.Reason);
            Assert.Contains("RC4", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Should_Report_UnsupportedScheme_When_Version_Unknown()
        {
            byte[] bogus = [0x63, 0x00, 0x63, 0x00, 0, 0, 0, 0];
            var ex = Assert.Throws<ExcelEncryptionException>(
                () => EncryptionDescriptor.Parse(bogus, ExcelReaderOptions.Default));
            Assert.Equal(ExcelEncryptionReason.UnsupportedScheme, ex.Reason);
        }

        // The spec permits CFB chaining; Excel never writes it and no fixture can be obtained,
        // so it is rejected honestly rather than half-supported.
        [Fact]
        public void Should_Report_UnsupportedScheme_When_Cipher_Chaining_Is_Cfb()
        {
            byte[] info = Info("agile-aes256-sha512.xlsx");
            string xml = Encoding.UTF8.GetString(info.AsSpan(8))
                .Replace("ChainingModeCBC", "ChainingModeCFB", StringComparison.Ordinal);
            byte[] patched = [.. info.AsSpan(0, 8), .. Encoding.UTF8.GetBytes(xml)];
            var ex = Assert.Throws<ExcelEncryptionException>(
                () => EncryptionDescriptor.Parse(patched, ExcelReaderOptions.Default));
            Assert.Equal(ExcelEncryptionReason.UnsupportedScheme, ex.Reason);
        }

        // AES-CBC (the only chaining mode this codebase performs — cipherChaining is restricted to
        // ChainingModeCBC above) has a fixed 16-byte block size. Any other declared blockSize (still
        // inside the old [1,64] bound) produces a wrong-sized IV in
        // AgileKeyDerivation.NormalizeToLength/SegmentIv, which throws a raw CryptographicException
        // out of `aes.IV = ...` — not one of this codebase's documented malformed-input exception
        // types — unless it's rejected here first.
        [Fact]
        public void Should_Report_UnsupportedScheme_When_BlockSize_Is_Not_Sixteen()
        {
            byte[] info = Info("agile-aes256-sha512.xlsx");
            string xml = Encoding.UTF8.GetString(info.AsSpan(8))
                .Replace("blockSize=\"16\"", "blockSize=\"15\"", StringComparison.Ordinal);
            byte[] patched = [.. info.AsSpan(0, 8), .. Encoding.UTF8.GetBytes(xml)];
            var ex = Assert.Throws<ExcelEncryptionException>(
                () => EncryptionDescriptor.Parse(patched, ExcelReaderOptions.Default));
            Assert.Equal(ExcelEncryptionReason.UnsupportedScheme, ex.Reason);
        }

        [Fact]
        public void Should_Report_UnsupportedScheme_When_KeyBits_Not_Allowed()
        {
            byte[] info = Info("agile-aes256-sha512.xlsx");
            string xml = Encoding.UTF8.GetString(info.AsSpan(8))
                .Replace("keyBits=\"256\"", "keyBits=\"777\"", StringComparison.Ordinal);
            byte[] patched = [.. info.AsSpan(0, 8), .. Encoding.UTF8.GetBytes(xml)];
            var ex = Assert.Throws<ExcelEncryptionException>(
                () => EncryptionDescriptor.Parse(patched, ExcelReaderOptions.Default));
            Assert.Equal(ExcelEncryptionReason.UnsupportedScheme, ex.Reason);
        }

        // The descriptor is parsed BEFORE any password check, so an XXE here would be reachable
        // pre-authentication on wholly untrusted input.
        [Fact]
        public void Should_Not_Resolve_External_Entities_When_Descriptor_Contains_Doctype()
        {
            string xml = """
                <?xml version="1.0"?>
                <!DOCTYPE encryption [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>
                <encryption><keyData saltValue="&xxe;"/></encryption>
                """;
            byte[] payload = [0x04, 0x00, 0x04, 0x00, 0x40, 0, 0, 0, .. Encoding.UTF8.GetBytes(xml)];
            Exception? ex = Record.Exception(
                () => EncryptionDescriptor.Parse(payload, ExcelReaderOptions.Default));
            Assert.NotNull(ex);
            Assert.True(ex is ExcelEncryptionException or InvalidDataException,
                $"DTD must be rejected outright, got {ex.GetType().Name}: {ex.Message}");
        }

        [Fact]
        public void Should_Throw_When_Truncated()
        {
            var ex = Assert.Throws<InvalidDataException>(
                () => EncryptionDescriptor.Parse([0x04, 0x00], ExcelReaderOptions.Default));
            Assert.NotNull(ex);
        }
    }
}
