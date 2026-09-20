using System.Buffers;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class ChainedWorkbookStreamTests
    {
        private const int SectorSize = 512;
        private const int HeaderSize = 512;
        private const int FileSectors = 3;

        private static (byte[] Buffer, int[] Chain) BuildFragmented(int firstSector = 2)
        {
            byte[] buffer = new byte[HeaderSize + (FileSectors * SectorSize)];
            int[] chain = ArrayPool<int>.Shared.Rent(2);
            chain[0] = firstSector;
            chain[1] = 0;
            return (buffer, chain);
        }

        private static int FileOffset(int sector, int within)
        {
            return HeaderSize + (sector * SectorSize) + within;
        }

        [Fact]
        public void ReadsARecordThatStraddlesAChainDiscontinuity()
        {
            (byte[] buffer, int[] chain) = BuildFragmented();

            buffer[FileOffset(2, 504)] = 0x03;
            buffer[FileOffset(2, 505)] = 0x02;
            buffer[FileOffset(2, 506)] = 8;
            buffer[FileOffset(2, 507)] = 0;
            for (int i = 0; i < 4; i++)
            {
                buffer[FileOffset(2, 508 + i)] = (byte)(i + 1);
                buffer[FileOffset(0, i)] = (byte)(i + 5);
            }

            using WorkbookStream wb = WorkbookStream.Chained(buffer, chain, chainLength: 2, SectorSize, length: 2 * SectorSize);
            using BiffCursor cursor = wb.OpenCursor();
            cursor.Position = 504;

            Assert.True(cursor.TryReadRecord(out int id, out ReadOnlySpan<byte> data));
            Assert.Equal(0x0203, id);
            Assert.Equal<byte>([1, 2, 3, 4, 5, 6, 7, 8], data.ToArray());
            Assert.Equal(516, cursor.Position);
        }

        [Fact]
        public void ReadsRecordsWhollyInsideOneSectorFromTheCachedRun()
        {
            (byte[] buffer, int[] chain) = BuildFragmented();

            buffer[FileOffset(2, 0)] = 0x03;
            buffer[FileOffset(2, 1)] = 0x02;
            buffer[FileOffset(2, 2)] = 2;
            buffer[FileOffset(2, 4)] = 0xAA;
            buffer[FileOffset(2, 5)] = 0xBB;
            buffer[FileOffset(2, 6)] = 0x05;
            buffer[FileOffset(2, 7)] = 0x02;
            buffer[FileOffset(2, 8)] = 1;
            buffer[FileOffset(2, 10)] = 0xCC;

            using WorkbookStream wb = WorkbookStream.Chained(buffer, chain, chainLength: 2, SectorSize, length: 2 * SectorSize);
            using BiffCursor cursor = wb.OpenCursor();

            Assert.True(cursor.TryReadRecord(out int first, out ReadOnlySpan<byte> firstData));
            Assert.Equal(0x0203, first);
            Assert.Equal<byte>([0xAA, 0xBB], firstData.ToArray());

            Assert.True(cursor.TryReadRecord(out int second, out ReadOnlySpan<byte> secondData));
            Assert.Equal(0x0205, second);
            Assert.Equal<byte>([0xCC], secondData.ToArray());
        }

        [Fact]
        public void ThrowsInvalidDataWhenAChainEntryPointsPastTheBuffer()
        {
            (byte[] buffer, int[] chain) = BuildFragmented(firstSector: 99);

            using WorkbookStream wb = WorkbookStream.Chained(buffer, chain, chainLength: 2, SectorSize, length: 2 * SectorSize);
            using BiffCursor cursor = wb.OpenCursor();

            Assert.Throws<InvalidDataException>(() => cursor.TryReadRecord(out _, out _));
        }
    }
}
