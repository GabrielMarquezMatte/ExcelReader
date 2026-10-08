using System.Buffers.Binary;
using System.IO.Compression;
using static ExcelReader.Tests.Reader.Zip.InflateTestSupport;

namespace ExcelReader.Tests.Reader.Zip
{
    public class InflateDifferentialTests
    {
        public static TheoryData<DataShape, CompressionLevel> ShapesAndLevels()
        {
            var data = new TheoryData<DataShape, CompressionLevel>();
            foreach (DataShape shape in Enum.GetValues<DataShape>())
            {
                foreach (CompressionLevel level in Enum.GetValues<CompressionLevel>())
                {
                    data.Add(shape, level);
                }
            }
            return data;
        }

        [Theory]
        [MemberData(nameof(ShapesAndLevels))]
        public void DeflateStreamOutputRoundTrips(DataShape shape, CompressionLevel level)
        {
            int[] sizes = [1, 2, 7, 100, 1000, 32768, 65536, 70_000, 300_000, 1_500_000];
            foreach (int size in sizes)
            {
                byte[] raw = Generate(shape, size, seed: size);
                byte[] stream = Deflate(raw, level);
                Assert.NotEmpty(stream);

                Assert.True(raw.AsSpan().SequenceEqual(Inflate(stream, readSize: 64 * 1024)), $"size {size}");
            }
        }

        [Theory]
        [InlineData("sample.xlsx")]
        [InlineData("RealExcel.xlsb")]
        public void EveryDeflateEntryOfAProducerWrittenWorkbookMatchesDeflateStream(string fileName)
        {
            byte[] zip = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "data", fileName));
            int compared = 0;

            foreach (byte[] stream in DeflateEntries(zip))
            {
                using var reference = new MemoryStream();
                using (var deflate = new DeflateStream(new MemoryStream(stream), CompressionMode.Decompress))
                {
                    deflate.CopyTo(reference);
                }

                Assert.True(reference.ToArray().AsSpan().SequenceEqual(Inflate(stream)), $"entry {compared} of {fileName}");
                compared++;
            }

            Assert.True(compared >= 5, $"{fileName} yielded only {compared} deflate entries");
        }

        private static List<byte[]> DeflateEntries(byte[] zip)
        {
            int end = zip.AsSpan().LastIndexOf("PK\x05\x06"u8);
            Assert.True(end >= 0, "no end-of-central-directory record");
            int count = BinaryPrimitives.ReadUInt16LittleEndian(zip.AsSpan(end + 10));
            int position = (int)BinaryPrimitives.ReadUInt32LittleEndian(zip.AsSpan(end + 16));
            var entries = new List<byte[]>();
            for (int i = 0; i < count; i++)
            {
                ReadOnlySpan<byte> header = zip.AsSpan(position);
                int method = BinaryPrimitives.ReadUInt16LittleEndian(header[10..]);
                int compressedSize = (int)BinaryPrimitives.ReadUInt32LittleEndian(header[20..]);
                int nameLength = BinaryPrimitives.ReadUInt16LittleEndian(header[28..]);
                int extraLength = BinaryPrimitives.ReadUInt16LittleEndian(header[30..]);
                int commentLength = BinaryPrimitives.ReadUInt16LittleEndian(header[32..]);
                int local = (int)BinaryPrimitives.ReadUInt32LittleEndian(header[42..]);
                if (method == 8)
                {
                    ReadOnlySpan<byte> localHeader = zip.AsSpan(local);
                    int data = local + 30 + BinaryPrimitives.ReadUInt16LittleEndian(localHeader[26..]) + BinaryPrimitives.ReadUInt16LittleEndian(localHeader[28..]);
                    entries.Add(zip.AsSpan(data, compressedSize).ToArray());
                }
                position += 46 + nameLength + extraLength + commentLength;
            }
            return entries;
        }
    }
}
