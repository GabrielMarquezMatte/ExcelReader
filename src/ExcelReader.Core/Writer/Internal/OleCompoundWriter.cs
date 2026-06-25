using System.Buffers.Binary;

namespace ExcelReader.Core.Writer.Internal
{
    // Wraps a Workbook BIFF stream in a minimal OLE/CFB container — the inverse of the
    // XlsCompoundFile reader. Layout: header(512) + FAT sectors + directory sector + workbook
    // sectors. The workbook is stored as a regular stream (size padded to >= the mini cutoff so
    // the reader never takes the mini-stream path). Trailing pad bytes sit after the EOF record,
    // which the reader stops at, so they are never parsed.
    internal static class OleCompoundWriter
    {
        private const int HeaderSize = 512;
        private const int SectorSize = 512;
        private const int MiniCutoff = 4096;
        private const int EndOfChain = unchecked((int)0xFFFFFFFE);
        private const int FatSectorMarker = unchecked((int)0xFFFFFFFD);
        private const int FreeSector = unchecked((int)0xFFFFFFFF);
        private const int FatEntriesPerSector = SectorSize / 4; // 128
        private const int MaxHeaderDifat = (HeaderSize - 0x4C) / 4; // 109

        // Writes the OLE container around a workbook stream of exactly workbookSize bytes. The body
        // is streamed by writeBody directly to destination — no combined buffer is materialized.
        internal static async ValueTask WriteAsync(Stream destination, int workbookSize, Func<Stream, CancellationToken, ValueTask> writeBody, CancellationToken ct)
        {
            // Pad to a whole number of sectors and keep the stored size at/above the mini cutoff
            // so the reader treats it as a regular stream.
            int storedSize = Math.Max(RoundUp(workbookSize, SectorSize), MiniCutoff);
            int workbookSectors = storedSize / SectorSize;

            // The FAT must cover its own sectors plus the directory and workbook sectors. Each FAT
            // sector holds 128 entries but also consumes one, hence the divide by 127 (not 128).
            int fatSectorCount = (workbookSectors + 1 + 126) / 127;
            if (fatSectorCount > MaxHeaderDifat)
            {
                throw new NotSupportedException("The workbook is too large for the simple OLE writer (would need DIFAT sectors).");
            }

            int directorySector = fatSectorCount;
            int workbookStart = directorySector + 1;

            byte[] header = BuildHeader(fatSectorCount, directorySector);
            byte[] fat = BuildFat(fatSectorCount, directorySector, workbookStart, workbookSectors);
            byte[] directory = BuildDirectory(workbookStart, storedSize);

            await destination.WriteAsync(header, ct).ConfigureAwait(false);
            await destination.WriteAsync(fat, ct).ConfigureAwait(false);
            await destination.WriteAsync(directory, ct).ConfigureAwait(false);
            await writeBody(destination, ct).ConfigureAwait(false);
            int padding = storedSize - workbookSize;
            if (padding > 0)
            {
                await destination.WriteAsync(new byte[padding], ct).ConfigureAwait(false);
            }
        }

        private static byte[] BuildHeader(int fatSectorCount, int directorySector)
        {
            byte[] header = new byte[HeaderSize];
            ReadOnlySpan<byte> signature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
            signature.CopyTo(header);
            WriteU16(header, 0x18, 0x003E); // minor version
            WriteU16(header, 0x1A, 0x0003); // major version
            WriteU16(header, 0x1C, 0xFFFE); // byte-order mark
            WriteU16(header, 0x1E, 9);      // sector shift -> 512
            WriteU16(header, 0x20, 6);      // mini-sector shift -> 64
            WriteI32(header, 0x2C, fatSectorCount);
            WriteI32(header, 0x30, directorySector);
            WriteI32(header, 0x38, MiniCutoff);
            WriteI32(header, 0x3C, EndOfChain); // first mini-FAT sector (none)
            WriteI32(header, 0x40, 0);          // mini-FAT sector count
            WriteI32(header, 0x44, EndOfChain); // first DIFAT sector (none)
            WriteI32(header, 0x48, 0);          // DIFAT sector count

            for (int i = 0x4C; i < HeaderSize; i += 4)
            {
                WriteI32(header, i, FreeSector);
            }
            for (int i = 0; i < fatSectorCount; i++)
            {
                WriteI32(header, 0x4C + (i * 4), i); // FAT sectors occupy sectors 0..fatSectorCount-1
            }
            return header;
        }

        private static byte[] BuildFat(int fatSectorCount, int directorySector, int workbookStart, int workbookSectors)
        {
            byte[] fat = new byte[fatSectorCount * FatEntriesPerSector * 4];
            fat.AsSpan().Fill(0xFF); // FreeSector everywhere by default

            for (int i = 0; i < fatSectorCount; i++)
            {
                WriteI32(fat, i * 4, FatSectorMarker);
            }
            WriteI32(fat, directorySector * 4, EndOfChain);
            for (int i = 0; i < workbookSectors; i++)
            {
                int sector = workbookStart + i;
                int next = i == workbookSectors - 1 ? EndOfChain : sector + 1;
                WriteI32(fat, sector * 4, next);
            }
            return fat;
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
