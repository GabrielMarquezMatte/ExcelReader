using System.Buffers.Binary;

namespace ExcelReader.Core.Writer.Internal
{
    // Wraps a Workbook BIFF stream in a minimal OLE/CFB container — the inverse of the
    // XlsCompoundFile reader. Layout: header(512) + FAT sectors + DIFAT sectors (if needed) +
    // directory sector + workbook sectors. The workbook is stored as a regular stream (size padded
    // to >= the mini cutoff so the reader never takes the mini-stream path). Trailing pad bytes sit
    // after the EOF record, which the reader stops at, so they are never parsed.
    internal static class OleCompoundWriter
    {
        private const int HeaderSize = 512;
        private const int SectorSize = 512;
        private const int MiniCutoff = 4096;
        private const int EndOfChain = unchecked((int)0xFFFFFFFE);
        private const int FatSectorMarker = unchecked((int)0xFFFFFFFD);
        private const int DifatSectorMarker = unchecked((int)0xFFFFFFFC);
        private const int FreeSector = unchecked((int)0xFFFFFFFF);
        private const int FatEntriesPerSector = SectorSize / 4;            // 128
        private const int MaxHeaderDifat = (HeaderSize - 0x4C) / 4;       // 109
        private const int DifatEntriesPerSector = FatEntriesPerSector - 1; // 127 (last slot = next DIFAT)

        // Writes the OLE container around a workbook stream of exactly workbookSize bytes. The body
        // is streamed by writeBody directly to destination — no combined buffer is materialized.
        internal static async ValueTask WriteAsync(Stream destination, int workbookSize, Func<Stream, CancellationToken, ValueTask> writeBody, CancellationToken ct)
        {
            // Pad to a whole number of sectors and keep the stored size at/above the mini cutoff
            // so the reader treats it as a regular stream.
            int storedSize = Math.Max(RoundUp(workbookSize, SectorSize), MiniCutoff);
            int workbookSectors = storedSize / SectorSize;

            (int fatCount, int difatCount) = ComputeSectorCounts(workbookSectors);

            int firstDifatSector = fatCount; // DIFAT sectors follow FAT sectors
            int directorySector = fatCount + difatCount;
            int workbookStart = directorySector + 1;

            byte[] header = BuildHeader(fatCount, difatCount, difatCount > 0 ? firstDifatSector : EndOfChain, directorySector);
            byte[] fat = BuildFat(fatCount, difatCount, directorySector, workbookStart, workbookSectors);
            byte[] directory = BuildDirectory(workbookStart, storedSize);

            await destination.WriteAsync(header, ct).ConfigureAwait(false);
            await destination.WriteAsync(fat, ct).ConfigureAwait(false);
            if (difatCount > 0)
            {
                await destination.WriteAsync(BuildDifat(fatCount, difatCount, firstDifatSector), ct).ConfigureAwait(false);
            }
            await destination.WriteAsync(directory, ct).ConfigureAwait(false);
            await writeBody(destination, ct).ConfigureAwait(false);
            int padding = storedSize - workbookSize;
            if (padding > 0)
            {
                await destination.WriteAsync(new byte[padding], ct).ConfigureAwait(false);
            }
        }

        // Iteratively resolves the circular dependency: more workbook data needs more FAT sectors,
        // more FAT sectors may require DIFAT sectors, and DIFAT sectors increase the total sector
        // count, which may require yet more FAT sectors. Converges in ≤ 2 iterations in practice.
        private static (int fatCount, int difatCount) ComputeSectorCounts(int workbookSectors)
        {
            int fat = CeilingDiv(workbookSectors + 1, 127); // seed: 1 dir sector, no DIFAT overhead
            for (int i = 0; i < 4; i++)
            {
                int difat = fat <= MaxHeaderDifat ? 0 : CeilingDiv(fat - MaxHeaderDifat, DifatEntriesPerSector);
                fat = CeilingDiv(fat + difat + 1 + workbookSectors, FatEntriesPerSector);
            }
            int finalDifat = fat <= MaxHeaderDifat ? 0 : CeilingDiv(fat - MaxHeaderDifat, DifatEntriesPerSector);
            return (fat, finalDifat);
        }

        private static byte[] BuildHeader(int fatCount, int difatCount, int firstDifatSector, int directorySector)
        {
            byte[] header = new byte[HeaderSize];
            ReadOnlySpan<byte> signature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
            signature.CopyTo(header);
            WriteU16(header, 0x18, 0x003E); // minor version
            WriteU16(header, 0x1A, 0x0003); // major version
            WriteU16(header, 0x1C, 0xFFFE); // byte-order mark
            WriteU16(header, 0x1E, 9);      // sector shift -> 512
            WriteU16(header, 0x20, 6);      // mini-sector shift -> 64
            WriteI32(header, 0x2C, fatCount);
            WriteI32(header, 0x30, directorySector);
            WriteI32(header, 0x38, MiniCutoff);
            WriteI32(header, 0x3C, EndOfChain); // first mini-FAT sector (none)
            WriteI32(header, 0x40, 0);          // mini-FAT sector count
            WriteI32(header, 0x44, firstDifatSector);
            WriteI32(header, 0x48, difatCount);

            for (int i = 0x4C; i < HeaderSize; i += 4)
            {
                WriteI32(header, i, FreeSector);
            }
            int headerFatCount = Math.Min(fatCount, MaxHeaderDifat);
            for (int i = 0; i < headerFatCount; i++)
            {
                WriteI32(header, 0x4C + (i * 4), i); // FAT sectors occupy sectors 0..fatCount-1
            }
            return header;
        }

        private static byte[] BuildFat(int fatCount, int difatCount, int directorySector, int workbookStart, int workbookSectors)
        {
            byte[] fat = new byte[fatCount * FatEntriesPerSector * 4];
            fat.AsSpan().Fill(0xFF); // FreeSector everywhere by default

            for (int i = 0; i < fatCount; i++)
                WriteI32(fat, i * 4, FatSectorMarker);
            for (int i = 0; i < difatCount; i++)
                WriteI32(fat, (fatCount + i) * 4, DifatSectorMarker);
            WriteI32(fat, directorySector * 4, EndOfChain);
            for (int i = 0; i < workbookSectors; i++)
            {
                int sector = workbookStart + i;
                int next = (i == workbookSectors - 1) ? EndOfChain : (sector + 1);
                WriteI32(fat, sector * 4, next);
            }
            return fat;
        }

        // Each DIFAT sector: 127 FAT-sector indices (FreeSector padding if fewer) + next-DIFAT pointer.
        private static byte[] BuildDifat(int fatCount, int difatCount, int firstDifatSector)
        {
            byte[] difat = new byte[difatCount * SectorSize];
            for (int d = 0; d < difatCount; d++)
            {
                int byteBase = d * SectorSize;
                int fatBase = MaxHeaderDifat + (d * DifatEntriesPerSector);
                for (int j = 0; j < DifatEntriesPerSector; j++)
                {
                    int fatIdx = fatBase + j;
                    WriteI32(difat, byteBase + (j * 4), fatIdx < fatCount ? fatIdx : FreeSector);
                }
                int next = (d < difatCount - 1) ? firstDifatSector + d + 1 : EndOfChain;
                WriteI32(difat, byteBase + (DifatEntriesPerSector * 4), next);
            }
            return difat;
        }

        private static byte[] BuildDirectory(int workbookStart, int workbookSize)
        {
            byte[] directory = new byte[SectorSize];
            WriteDirectoryEntry(directory.AsSpan(0, 128), "Root Entry", objectType: 5, startSector: EndOfChain, size: 0, child: 1);
            WriteDirectoryEntry(directory.AsSpan(128, 128), "Workbook", objectType: 2, startSector: workbookStart, size: workbookSize, child: EndOfChain);
            return directory;
        }

        private static void WriteDirectoryEntry(Span<byte> entry, string name, byte objectType, int startSector, long size, int child)
        {
            System.Text.Encoding.Unicode.GetBytes(name + '\0').CopyTo(entry);
            WriteU16(entry, 64, (ushort)((name.Length + 1) * 2)); // name byte length incl terminator
            entry[66] = objectType;
            entry[67] = 1; // color = black
            WriteI32(entry, 68, EndOfChain); // left sibling
            WriteI32(entry, 72, EndOfChain); // right sibling
            WriteI32(entry, 76, child);
            WriteI32(entry, 116, startSector);
            BinaryPrimitives.WriteInt64LittleEndian(entry[120..], size);
        }

        private static int CeilingDiv(int n, int d) => (n + d - 1) / d;

        private static int RoundUp(int value, int multiple)
        {
            return (value + multiple - 1) / multiple * multiple;
        }

        private static void WriteU16(Span<byte> dest, int offset, ushort value)
        {
            BinaryPrimitives.WriteUInt16LittleEndian(dest.Slice(offset, 2), value);
        }

        private static void WriteI32(Span<byte> dest, int offset, int value)
        {
            BinaryPrimitives.WriteInt32LittleEndian(dest.Slice(offset, 4), value);
        }
    }
}
