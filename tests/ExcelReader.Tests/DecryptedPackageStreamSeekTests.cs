using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class DecryptedPackageStreamSeekTests
    {
        private const string Fixture = "agile-aes256-sha512.xlsx";

        private static DecryptedPackageStream Open(out CfbContainer cfb)
        {
            var fs = File.OpenRead(EncryptedFixtures.Path_(Fixture));
            cfb = CfbContainer.Parse(fs, ownsSource: true, ExcelReaderOptions.Default);
            byte[] info = cfb.ReadStream("EncryptionInfo", 64 * 1024);
            var options = ExcelReaderOptions.Default with { Password = EncryptedFixtures.Password };
            return DecryptedPackageStream.Create(cfb, EncryptionDescriptor.Parse(info, options), options);
        }

        // Every offset in the first three segments, plus the segment boundaries themselves: a
        // one-segment-off cache or a bad in-segment offset shows up as a mismatch at exactly one
        // of these, which pinpoints the bug.
        [Fact]
        public void Should_Match_Oracle_When_Seeking_To_Every_Offset_Near_Boundaries()
        {
            byte[] expected = EncryptedFixtures.PlainBytes(Fixture);
            using DecryptedPackageStream stream = Open(out CfbContainer cfb);
            using (cfb)
            {
                foreach (int origin in new[] { 0, 1, 4095, 4096, 4097, 8191, 8192, 8193 })
                {
                    if (origin >= expected.Length)
                    {
                        continue;
                    }
                    stream.Position = origin;
                    int take = Math.Min(300, expected.Length - origin);
                    byte[] actual = new byte[take];
                    stream.ReadExactly(actual);
                    Assert.Equal(expected.AsSpan(origin, take).ToArray(), actual);
                }
            }
        }

        // ZipArchive's first move is to find the end-of-central-directory record at the tail.
        [Fact]
        public void Should_Match_Oracle_When_Seeking_From_End()
        {
            byte[] expected = EncryptedFixtures.PlainBytes(Fixture);
            using DecryptedPackageStream stream = Open(out CfbContainer cfb);
            using (cfb)
            {
                stream.Seek(-22, SeekOrigin.End);
                byte[] actual = new byte[22];
                stream.ReadExactly(actual);
                Assert.Equal(expected[^22..], actual);
            }
        }

        // Backwards seeks must invalidate nothing incorrectly: re-reading an earlier segment after
        // a later one is exactly what ZipArchive does after reading the central directory.
        [Fact]
        public void Should_Match_Oracle_When_Seeking_Backwards_After_Reading_Tail()
        {
            byte[] expected = EncryptedFixtures.PlainBytes(Fixture);
            using DecryptedPackageStream stream = Open(out CfbContainer cfb);
            using (cfb)
            {
                stream.Seek(-22, SeekOrigin.End);
                stream.ReadExactly(new byte[22]);

                stream.Position = 0;
                byte[] head = new byte[64];
                stream.ReadExactly(head);
                Assert.Equal(expected.AsSpan(0, 64).ToArray(), head);
            }
        }

        [Fact]
        public void Should_Track_Position_When_Seeking_With_Each_Origin()
        {
            using DecryptedPackageStream stream = Open(out CfbContainer cfb);
            using (cfb)
            {
                Assert.Equal(100, stream.Seek(100, SeekOrigin.Begin));
                Assert.Equal(150, stream.Seek(50, SeekOrigin.Current));
                Assert.Equal(stream.Length - 10, stream.Seek(-10, SeekOrigin.End));
                Assert.Equal(stream.Length - 10, stream.Position);
            }
        }

        [Fact]
        public void Should_Throw_When_Seeking_Before_Start()
        {
            using DecryptedPackageStream stream = Open(out CfbContainer cfb);
            using (cfb)
            {
                Assert.Throws<ArgumentOutOfRangeException>(() => stream.Seek(-1, SeekOrigin.Begin));
                Assert.Throws<ArgumentOutOfRangeException>(() => stream.Position = -1);
            }
        }

        [Fact]
        public void Should_Return_Zero_When_Reading_At_End()
        {
            using DecryptedPackageStream stream = Open(out CfbContainer cfb);
            using (cfb)
            {
                stream.Position = stream.Length;
                Assert.Equal(0, stream.Read(new byte[16]));
            }
        }
    }
}
