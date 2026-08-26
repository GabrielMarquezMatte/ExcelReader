using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class DecryptedPackageStreamTests
    {
        private static DecryptedPackageStream Open(string fixture, out CfbContainer cfb)
        {
            var fs = File.OpenRead(EncryptedFixtures.Path_(fixture));
            cfb = CfbContainer.Parse(fs, ownsSource: true, ExcelReaderOptions.Default);
            byte[] info = cfb.ReadStream("EncryptionInfo", 64 * 1024);
            var options = ExcelReaderOptions.Default with { Password = EncryptedFixtures.Password };
            EncryptionDescriptor d = EncryptionDescriptor.Parse(info, options);
            return DecryptedPackageStream.Create(cfb, d, options);
        }

        public static TheoryData<string> Fixtures()
        {
            var data = new TheoryData<string>();
            foreach (string name in EncryptedFixtures.All)
            {
                data.Add(name);
            }
            return data;
        }

        // The oracle: msoffcrypto-tool is an independent implementation, so byte-exact agreement is
        // the strongest correctness signal available with no writer to round-trip against.
        [Theory]
        [MemberData(nameof(Fixtures))]
        public void Should_Match_Oracle_When_Read_Sequentially(string fixture)
        {
            using DecryptedPackageStream stream = Open(fixture, out CfbContainer cfb);
            using (cfb)
            {
                byte[] expected = EncryptedFixtures.PlainBytes(fixture);
                Assert.Equal(expected.Length, stream.Length);
                byte[] actual = new byte[stream.Length];
                stream.ReadExactly(actual);
                Assert.Equal(expected, actual);
            }
        }

        [Theory]
        [MemberData(nameof(Fixtures))]
        public async Task Should_Match_Oracle_When_Read_Async(string fixture)
        {
            using DecryptedPackageStream stream = Open(fixture, out CfbContainer cfb);
            using (cfb)
            {
                byte[] expected = EncryptedFixtures.PlainBytes(fixture);
                byte[] actual = new byte[stream.Length];
                await stream.ReadExactlyAsync(actual, TestContext.Current.CancellationToken);
                Assert.Equal(expected, actual);
            }
        }

        // Reading one byte at a time crosses every segment boundary the hard way; a segment cache
        // that refills incorrectly at the boundary fails here and nowhere else.
        [Fact]
        public void Should_Match_Oracle_When_Read_One_Byte_At_A_Time()
        {
            using DecryptedPackageStream stream = Open("agile-aes256-sha512.xlsx", out CfbContainer cfb);
            using (cfb)
            {
                byte[] expected = EncryptedFixtures.PlainBytes("agile-aes256-sha512.xlsx");
                foreach (byte expectedByte in expected)
                {
                    int b = stream.ReadByte();
                    Assert.Equal(expectedByte, (byte)b);
                }
                Assert.Equal(-1, stream.ReadByte());
            }
        }

        // The final segment is short: the ciphertext is padded to a 16-byte boundary while the
        // plaintext is truncated to the declared length. Off-by-one here yields trailing garbage
        // that ZipArchive would reject much later with a confusing message.
        [Theory]
        [MemberData(nameof(Fixtures))]
        public void Should_Stop_At_Declared_Length_When_Ciphertext_Is_Padded(string fixture)
        {
            using DecryptedPackageStream stream = Open(fixture, out CfbContainer cfb);
            using (cfb)
            {
                byte[] buffer = new byte[stream.Length + 64];
                int total = 0, read;
                while ((read = stream.Read(buffer.AsSpan(total))) > 0)
                {
                    total += read;
                }
                Assert.Equal(stream.Length, total);
            }
        }

        [Fact]
        public void Should_Reject_Writes()
        {
            using DecryptedPackageStream stream = Open("agile-aes256-sha512.xlsx", out CfbContainer cfb);
            using (cfb)
            {
                Assert.False(stream.CanWrite);
                Assert.Throws<NotSupportedException>(() => stream.Write([1, 2, 3]));
                Assert.Throws<NotSupportedException>(() => stream.SetLength(0));
            }
        }
    }
}
