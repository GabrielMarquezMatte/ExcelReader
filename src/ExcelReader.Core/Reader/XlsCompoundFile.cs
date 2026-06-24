using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Reader
{
    [ExcludeFromCodeCoverage(Justification = "Covered through XlsReader integration tests; most uncovered paths are corrupt-OLE guard rails.")]
    internal sealed class XlsCompoundFile
    {
        private const int HeaderSize = 512;
        private const int EndOfChain = unchecked((int)0xFFFFFFFE);
        private const int FatSector = unchecked((int)0xFFFFFFFD);
        private const int FreeSector = unchecked((int)0xFFFFFFFF);
        private static readonly byte[] _signature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

        private readonly byte[] _bytes;
        private readonly int _sectorSize;
        private readonly int _miniSectorSize;
        private readonly int _miniCutoff;
        private readonly int[] _fat;
        private readonly int[] _miniFat;
        private readonly DirectoryEntry[] _entries;
        private readonly byte[] _miniStream;

        private XlsCompoundFile(
            byte[] bytes,
            int sectorSize,
            int miniSectorSize,
            int miniCutoff,
            int[] fat,
            int[] miniFat,
            DirectoryEntry[] entries,
            byte[] miniStream)
        {
            _bytes = bytes;
            _sectorSize = sectorSize;
            _miniSectorSize = miniSectorSize;
            _miniCutoff = miniCutoff;
            _fat = fat;
            _miniFat = miniFat;
            _entries = entries;
            _miniStream = miniStream;
        }

        internal static XlsCompoundFile Open(Stream stream, bool leaveOpen)
        {
            try
            {
                if (TryGetExactLength(stream, out int length))
                {
                    byte[] buffer = new byte[length];
                    stream.ReadExactly(buffer);
                    return Open(buffer);
                }
                using MemoryStream ms = new();
                stream.CopyTo(ms);
                return Open(ms.ToArray());
            }
            finally
            {
                if (!leaveOpen)
                {
                    stream.Dispose();
                }
            }
        }

        internal static async ValueTask<XlsCompoundFile> OpenAsync(Stream stream, bool leaveOpen, CancellationToken ct)
        {
            try
            {
                if (TryGetExactLength(stream, out int length))
                {
                    byte[] buffer = new byte[length];
                    await stream.ReadExactlyAsync(buffer, ct).ConfigureAwait(false);
                    return Open(buffer);
                }
                MemoryStream ms = new();
                await using (ms.ConfigureAwait(false))
                {
                    await stream.CopyToAsync(ms, ct).ConfigureAwait(false);
                    return Open(ms.ToArray());
                }
            }
            finally
            {
                if (!leaveOpen)
                {
                    await stream.DisposeAsync().ConfigureAwait(false);
                }
            }
        }

        private static bool TryGetExactLength(Stream stream, out int length)
        {
            length = 0;
            if (!stream.CanSeek)
            {
                return false;
            }
            long remaining = stream.Length - stream.Position;
            if (remaining is <= 0 or > int.MaxValue)
            {
                return false;
            }
            length = (int)remaining;
            return true;
        }

        private static XlsCompoundFile Open(byte[] bytes)
        {
            if (bytes.Length < HeaderSize || !bytes.AsSpan(0, _signature.Length).SequenceEqual(_signature))
            {
                throw new InvalidDataException("The stream is not an OLE compound document.");
            }

            int sectorSize = 1 << ReadU16(bytes, 0x1E);
            int miniSectorSize = 1 << ReadU16(bytes, 0x20);
            int fatSectorCount = ReadI32(bytes, 0x2C);
            int firstDirectorySector = ReadI32(bytes, 0x30);
            int miniCutoff = ReadI32(bytes, 0x38);
            int firstMiniFatSector = ReadI32(bytes, 0x3C);
            int miniFatSectorCount = ReadI32(bytes, 0x40);
            int firstDifatSector = ReadI32(bytes, 0x44);
            int difatSectorCount = ReadI32(bytes, 0x48);

            if (sectorSize < HeaderSize || sectorSize > 4096 || miniSectorSize <= 0)
            {
                throw new InvalidDataException("Unsupported OLE sector size.");
            }

            int[] fatSectorIds = ReadDifat(bytes, sectorSize, fatSectorCount, firstDifatSector, difatSectorCount);
            int[] fat = ReadFat(bytes, sectorSize, fatSectorIds);
            byte[] directoryStream = ReadRegularStream(bytes, sectorSize, fat, firstDirectorySector, int.MaxValue);
            DirectoryEntry[] entries = ReadDirectory(directoryStream);
            if (entries.Length == 0)
            {
                throw new InvalidDataException("The OLE directory is empty.");
            }

            int[] miniFat = firstMiniFatSector >= 0 && miniFatSectorCount > 0
                ? ReadMiniFat(bytes, sectorSize, fat, firstMiniFatSector, miniFatSectorCount)
                : [];
            byte[] miniStream = entries[0].StartSector >= 0 && entries[0].Size > 0
                ? ReadRegularStream(bytes, sectorSize, fat, entries[0].StartSector, entries[0].Size)
                : [];

            return new XlsCompoundFile(bytes, sectorSize, miniSectorSize, miniCutoff, fat, miniFat, entries, miniStream);
        }

        internal byte[] ReadWorkbookStream()
        {
            foreach (ref readonly var entry in _entries.AsSpan())
            {
                if (entry.ObjectType == 2 &&
                    (entry.Name.Equals("Workbook", StringComparison.OrdinalIgnoreCase) ||
                     entry.Name.Equals("Book", StringComparison.OrdinalIgnoreCase)))
                {
                    return ReadStream(entry);
                }
            }
            throw new InvalidDataException("The OLE document does not contain a Workbook stream.");
        }

        private byte[] ReadStream(DirectoryEntry entry)
        {
            if (entry.Size < _miniCutoff && entry.StartSector >= 0)
            {
                if (_miniFat.Length == 0 || _miniStream.Length == 0)
                {
                    throw new InvalidDataException("The OLE mini stream is missing.");
                }
                return ReadMiniStream(entry.StartSector, entry.Size);
            }
            return ReadRegularStream(_bytes, _sectorSize, _fat, entry.StartSector, entry.Size);
        }

        private byte[] ReadMiniStream(int startSector, long size)
        {
            byte[] result = new byte[(int)size];
            int sector = startSector;
            int written = 0;
            while (sector >= 0 && sector != EndOfChain && written < result.Length)
            {
                int offset = checked(sector * _miniSectorSize);
                if ((uint)offset >= (uint)_miniStream.Length)
                {
                    throw new InvalidDataException("Invalid OLE mini sector chain.");
                }
                int take = Math.Min(_miniSectorSize, result.Length - written);
                _miniStream.AsSpan(offset, take).CopyTo(result.AsSpan(written));
                written += take;
                if ((uint)sector >= (uint)_miniFat.Length)
                {
                    throw new InvalidDataException("Invalid OLE mini FAT chain.");
                }
                sector = _miniFat[sector];
            }
            return result;
        }

        private static int[] ReadDifat(ReadOnlySpan<byte> bytes, int sectorSize, int fatSectorCount, int firstDifatSector, int difatSectorCount)
        {
            int[] fatSectors = new int[fatSectorCount];
            int count = 0;
            for (int i = 0x4C; i < HeaderSize && count < fatSectors.Length; i += 4)
            {
                int sector = ReadI32(bytes, i);
                if (sector is >= 0 and not FreeSector)
                {
                    fatSectors[count++] = sector;
                }
            }

            int difat = firstDifatSector;
            for (int i = 0; i < difatSectorCount && difat >= 0 && count < fatSectors.Length; i++)
            {
                int offset = SectorOffset(difat, sectorSize, bytes.Length);
                int entries = (sectorSize / 4) - 1;
                for (int j = 0; j < entries && count < fatSectors.Length; j++)
                {
                    int sector = ReadI32(bytes, offset + (j * 4));
                    if (sector is >= 0 and not FreeSector)
                    {
                        fatSectors[count++] = sector;
                    }
                }
                difat = ReadI32(bytes, offset + (entries * 4));
            }

            if (count != fatSectors.Length)
            {
                throw new InvalidDataException("The OLE DIFAT is incomplete.");
            }
            return fatSectors;
        }

        private static int[] ReadFat(byte[] bytes, int sectorSize, int[] fatSectorIds)
        {
            int entriesPerSector = sectorSize / 4;
            int[] fat = new int[fatSectorIds.Length * entriesPerSector];
            int index = 0;
            foreach (int sector in fatSectorIds)
            {
                int offset = SectorOffset(sector, sectorSize, bytes.Length);
                for (int i = 0; i < entriesPerSector; i++)
                {
                    fat[index++] = ReadI32(bytes, offset + (i * 4));
                }
            }
            return fat;
        }

        private static int[] ReadMiniFat(byte[] bytes, int sectorSize, int[] fat, int firstSector, int sectorCount)
        {
            byte[] data = ReadRegularStream(bytes, sectorSize, fat, firstSector, checked(sectorCount * sectorSize));
            int[] miniFat = new int[data.Length / 4];
            for (int i = 0; i < miniFat.Length; i++)
            {
                miniFat[i] = ReadI32(data, i * 4);
            }
            return miniFat;
        }

        private static byte[] ReadRegularStream(byte[] bytes, int sectorSize, int[] fat, int startSector, long size)
        {
            if (startSector < 0)
            {
                return [];
            }

            // Known size (the common case: workbook, mini-FAT): fill an exact array
            // directly via spans, skipping the MemoryStream's internal buffer + ToArray copy.
            if (size != int.MaxValue)
            {
                byte[] result = new byte[(int)size];
                int sector = startSector;
                int written = 0;
                while (sector >= 0 && sector != EndOfChain && written < result.Length)
                {
                    int offset = SectorOffset(sector, sectorSize, bytes.Length);
                    int take = Math.Min(sectorSize, result.Length - written);
                    bytes.AsSpan(offset, take).CopyTo(result.AsSpan(written));
                    written += take;
                    sector = NextSector(fat, sector);
                }
                return result;
            }

            // Unknown size (directory stream): grow a MemoryStream.
            using MemoryStream ms = new();
            int dirSector = startSector;
            while (dirSector is >= 0 and not EndOfChain)
            {
                int offset = SectorOffset(dirSector, sectorSize, bytes.Length);
                ms.Write(bytes, offset, sectorSize);
                dirSector = NextSector(fat, dirSector);
            }
            return ms.ToArray();
        }

        private static int NextSector(int[] fat, int sector)
        {
            if ((uint)sector >= (uint)fat.Length)
            {
                throw new InvalidDataException("Invalid OLE FAT chain.");
            }
            int next = fat[sector];
            if (next is FatSector or FreeSector)
            {
                throw new InvalidDataException("Invalid OLE FAT sector reference.");
            }
            return next;
        }

        private static DirectoryEntry[] ReadDirectory(ReadOnlySpan<byte> bytes)
        {
            int count = bytes.Length / 128;
            DirectoryEntry[] entries = new DirectoryEntry[count];
            for (int i = 0; i < count; i++)
            {
                ReadOnlySpan<byte> entry = bytes.Slice(i * 128, 128);
                int nameBytes = ReadU16(entry, 64);
                string name = string.Empty;
                if (nameBytes >= 2)
                {
                    name = System.Text.Encoding.Unicode.GetString(entry[..(nameBytes - 2)]);
                }
                entries[i] = new DirectoryEntry(
                    name,
                    entry[66],
                    ReadI32(entry, 116),
                    BinaryPrimitives.ReadInt64LittleEndian(entry.Slice(120, 8)));
            }
            return entries;
        }

        private static int SectorOffset(int sector, int sectorSize, int length)
        {
            long offset = HeaderSize + ((long)sector * sectorSize);
            if (sector < 0 || offset < 0 || offset + sectorSize > length)
            {
                throw new InvalidDataException("Invalid OLE sector offset.");
            }
            return (int)offset;
        }

        private static ushort ReadU16(ReadOnlySpan<byte> src, int offset)
        {
            return BinaryPrimitives.ReadUInt16LittleEndian(src.Slice(offset, 2));
        }

        private static int ReadI32(ReadOnlySpan<byte> src, int offset)
        {
            return BinaryPrimitives.ReadInt32LittleEndian(src.Slice(offset, 4));
        }

        private readonly record struct DirectoryEntry(string Name, byte ObjectType, int StartSector, long Size);
    }
}
