using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class Biff12RecordReaderTests
    {
        [Fact]
        public void ReadsSequenceOfRecordsWithOneAndTwoByteFraming()
        {
            // id 5 / len 8, then id 159 (2-byte id) / len 0, then id 2 / len 130 (2-byte length).
            byte[] data =
            [
                .. Record(5, new byte[8]),
                .. Record(159, []),
                .. Record(2, Sequential(130)),
            ];

            var reader = new Biff12RecordReader(data);

            Assert.True(reader.TryReadRecord(out int id0, out var p0));
            Assert.Equal(5, id0);
            Assert.Equal(8, p0.Length);

            Assert.True(reader.TryReadRecord(out int id1, out var p1));
            Assert.Equal(159, id1);
            Assert.Equal(0, p1.Length);

            Assert.True(reader.TryReadRecord(out int id2, out var p2));
            Assert.Equal(2, id2);
            Assert.Equal(130, p2.Length);
            Assert.Equal(0, p2[0]);
            Assert.Equal(129, p2[129]); // 130th byte, 0-based value 129

            Assert.False(reader.TryReadRecord(out _, out _));
            Assert.Equal(data.Length, reader.Position);
        }

        [Fact]
        public void TruncatedRecordPayloadReturnsFalseWithoutAdvancing()
        {
            // id 5 declares len 8 but only 3 payload bytes are present.
            byte[] data = [.. Id(5), .. Length(8), 1, 2, 3];
            var reader = new Biff12RecordReader(data);

            Assert.False(reader.TryReadRecord(out _, out _));
            Assert.Equal(0, reader.Position);
        }

        [Fact]
        public void TruncatedLengthVarintReturnsFalse()
        {
            // id 5, then a length byte with the continuation bit set but no following byte.
            byte[] data = [.. Id(5), 0x80];
            var reader = new Biff12RecordReader(data);

            Assert.False(reader.TryReadRecord(out _, out _));
        }

        [Fact]
        public void EmptyDataReadsNothing()
        {
            var reader = new Biff12RecordReader([]);
            Assert.False(reader.TryReadRecord(out _, out _));
            Assert.Equal(0, reader.Position);
        }

        private static byte[] Record(int id, byte[] payload)
        {
            return [.. Id(id), .. Length(payload.Length), .. payload];
        }

        private static byte[] Id(int id)
        {
            if (id < 0x80)
            {
                return [(byte)id];
            }
            return [(byte)((id & 0x7F) | 0x80), (byte)(id >> 7)];
        }

        private static byte[] Length(int value)
        {
            List<byte> bytes = [];
            do
            {
                byte b = (byte)(value & 0x7F);
                value >>= 7;
                if (value != 0)
                {
                    b |= 0x80;
                }
                bytes.Add(b);
            }
            while (value != 0);
            return [.. bytes];
        }

        private static byte[] Sequential(int count)
        {
            byte[] bytes = new byte[count];
            for (int i = 0; i < count; i++)
            {
                bytes[i] = (byte)i;
            }
            return bytes;
        }
    }
}
