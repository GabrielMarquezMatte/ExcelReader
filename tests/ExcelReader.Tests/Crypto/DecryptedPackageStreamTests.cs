using System.Buffers.Binary;
using System.Text;
using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests.Crypto
{
    public class DecryptedPackageStreamTests
    {
        private static DecryptedPackageStream Open(string fixture, out CfbContainer cfb)
        {
            var fs = File.OpenRead(EncryptedFixtures.Path_(fixture));
            cfb = CfbContainer.Parse(fs, ownsSource: true, ExcelReaderOptions.Default);
            byte[] info = cfb.ReadStream("EncryptionInfo", 64 * 1024);
            var options = ExcelReaderOptions.Default with { Password = EncryptedFixtures.PasswordFor(fixture) };
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

        [Theory]
        [MemberData(nameof(Fixtures))]
        public void Should_Reject_NonBlockAligned_Ciphertext(string fixture)
        {
            byte[] raw = File.ReadAllBytes(EncryptedFixtures.Path_(fixture));

            byte[] namePattern = Encoding.Unicode.GetBytes("EncryptedPackage\0");
            int entryOffset = raw.AsSpan().IndexOf(namePattern);
            Assert.True(entryOffset >= 0, "Could not locate the EncryptedPackage directory entry.");
            int sizeFieldOffset = entryOffset + 120;
            int startSectorFieldOffset = entryOffset + 116;

            long originalSize = BinaryPrimitives.ReadInt64LittleEndian(raw.AsSpan(sizeFieldOffset, 8));
            int startSector = BinaryPrimitives.ReadInt32LittleEndian(raw.AsSpan(startSectorFieldOffset, 4));
            int sectorSize = 1 << BinaryPrimitives.ReadUInt16LittleEndian(raw.AsSpan(0x1E, 2));
            long dataOffset = ((long)startSector + 1) * sectorSize;

            const int PrefixSize = 8;
            long newSize = originalSize - 5;
            long newDeclared = newSize - PrefixSize;
            BinaryPrimitives.WriteInt64LittleEndian(raw.AsSpan(sizeFieldOffset, 8), newSize);
            BinaryPrimitives.WriteInt64LittleEndian(raw.AsSpan((int)dataOffset, 8), newDeclared);

            using var source = new MemoryStream(raw);
            using CfbContainer cfb = CfbContainer.Parse(source, ownsSource: true, ExcelReaderOptions.Default);
            byte[] info = cfb.ReadStream("EncryptionInfo", 64 * 1024);
            var options = ExcelReaderOptions.Default with { Password = EncryptedFixtures.PasswordFor(fixture) };
            EncryptionDescriptor d = EncryptionDescriptor.Parse(info, options);

            InvalidDataException ex = Assert.Throws<InvalidDataException>(() => DecryptedPackageStream.Create(cfb, d, options));
            Assert.Contains("block size", ex.Message, StringComparison.OrdinalIgnoreCase);
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
