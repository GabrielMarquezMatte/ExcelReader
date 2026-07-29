using System.Buffers;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    // The chained (whole-file-in-memory + FAT chain) WorkbookStream mode, driven directly rather than
    // through XlsReader: XlsWorkbookBuilder always emits a strictly sequential sector chain, so the
    // discontinuity handling below — run resolution and the sector-by-sector assembly a straddling
    // record falls back to — is unreachable from the builder-based fixtures.
    public class ChainedWorkbookStreamTests
    {
        private const int SectorSize = 512;
        private const int HeaderSize = 512;
        private const int FileSectors = 3;

        // Logical sector 0 -> file sector 2, logical sector 1 -> file sector 0. The chain is therefore
        // discontinuous at the boundary (0 != 2 + 1), so a record crossing it cannot be served as one
        // zero-copy slice.
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

            // Header at logical 504 (inside logical sector 0), so its 8 data bytes span logical
            // 508..515 — four bytes at the tail of file sector 2, four at the head of file sector 0.
            buffer[FileOffset(2, 504)] = 0x03;
            buffer[FileOffset(2, 505)] = 0x02; // id 0x0203
            buffer[FileOffset(2, 506)] = 8;
            buffer[FileOffset(2, 507)] = 0;    // len 8
            for (int i = 0; i < 4; i++)
            {
                buffer[FileOffset(2, 508 + i)] = (byte)(i + 1);  // logical 508..511 -> 1,2,3,4
                buffer[FileOffset(0, i)] = (byte)(i + 5);        // logical 512..515 -> 5,6,7,8
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

            // Two back-to-back records inside logical sector 0; the second must come from the run
            // cached by the first, not a re-resolve.
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

        // A crafted chain entry pointing past the buffer must surface as InvalidDataException, the same
        // type the streamed path raises for a corrupt container — not an IndexOutOfRange from slicing,
        // and never a read of whatever happens to sit past the workbook in the caller's buffer.
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
