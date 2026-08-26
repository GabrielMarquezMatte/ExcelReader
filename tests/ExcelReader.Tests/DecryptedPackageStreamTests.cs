using System.Buffers.Binary;
using System.Text;
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

        // EncryptedPackage's declared plaintext length and CFB entry size are both attacker-controlled
        // neither is required by the CFB/OLE format itself to keep the ciphertext block-aligned. This
        // patches a real fixture's on-disk EncryptedPackage entry (both its directory Size field and
        // its on-disk declared-length prefix) 5 bytes shorter than the real, valid ciphertext length —
        // real ciphertext is always a multiple of 16 (every segment is either a full 4096-byte segment
        // or the final segment padded to a 16-byte boundary), so subtracting 5 always lands on a
        // non-block-aligned total, reliably reproducing a crafted file's final segment landing mid-block
        // (the reviewer's concrete repro: an 11-byte trailing partial block) regardless of either
        // fixture's exact real length. Create must reject this with InvalidDataException, not let
        // AES's PaddingMode.None decryptor throw a raw ArgumentOutOfRangeException from deep inside
        // EnsureSegment.
        [Theory]
        [MemberData(nameof(Fixtures))]
        public void Should_Reject_NonBlockAligned_Ciphertext(string fixture)
        {
            byte[] raw = File.ReadAllBytes(EncryptedFixtures.Path_(fixture));

            // Locate the "EncryptedPackage" directory entry by its raw UTF-16LE name (fixed at the
            // start of its 128-byte record) - simpler than re-deriving CfbContainer's own directory
            // sector-chain walk, and just as reliable for a single, distinctive 36-byte pattern.
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
            var options = ExcelReaderOptions.Default with { Password = EncryptedFixtures.Password };
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
