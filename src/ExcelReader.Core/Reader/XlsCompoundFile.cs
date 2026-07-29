using System.Buffers;
using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using static ExcelReader.Core.Reader.Biff12;

namespace ExcelReader.Core.Reader
{
    // Parses the OLE/CFB container metadata (header, FAT, directory, mini-FAT/stream) by seeking
    // a seekable source, never materializing the whole file. The Workbook stream is then handed
    // back as a WorkbookStream that reads its sectors on demand. Non-seekable sources are buffered
    // into a MemoryStream first (rare fallback, same cost as before).
    [ExcludeFromCodeCoverage(Justification = "Covered through XlsReader integration tests; most uncovered paths are corrupt-OLE guard rails.")]
    internal static class XlsCompoundFile
    {
        private const int HeaderSize = 512;
        private const int EndOfChain = unchecked((int)0xFFFFFFFE);
        private const int FatSector = unchecked((int)0xFFFFFFFD);
        private const int FreeSector = unchecked((int)0xFFFFFFFF);
        internal static ReadOnlySpan<byte> Signature => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

        internal static WorkbookStream OpenWorkbook(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null)
        {
            (Stream source, bool ownsSource) = EnsureSeekable(stream, leaveOpen);
            try
            {
                return BuildWorkbook(source, ownsSource, options ?? ExcelReaderOptions.Default);
            }
            catch
            {
                if (ownsSource)
                {
                    source.Dispose();
                }
                throw;
            }
        }

        internal static WorkbookStream OpenWorkbook(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            using MemoryStream metadata = AsStream(data);
            return BuildWorkbook(metadata, ownsSource: false, options ?? ExcelReaderOptions.Default, memory: data);
        }

        internal static async ValueTask<WorkbookStream> OpenWorkbookAsync(Stream stream, bool leaveOpen, ExcelReaderOptions? options, CancellationToken ct)
        {
            (Stream source, bool ownsSource) = await EnsureSeekableAsync(stream, leaveOpen, ct).ConfigureAwait(false);
            try
            {
                return BuildWorkbook(source, ownsSource, options ?? ExcelReaderOptions.Default);
            }
            catch
            {
                if (ownsSource)
                {
                    await source.DisposeAsync().ConfigureAwait(false);
                }
                throw;
            }
        }

        private static (Stream Source, bool OwnsSource) EnsureSeekable(Stream stream, bool leaveOpen)
        {
            if (stream.CanSeek)
            {
                return (stream, !leaveOpen);
            }
            MemoryStream ms = new();
            stream.CopyTo(ms);
            if (!leaveOpen)
            {
                stream.Dispose();
            }
            ms.Position = 0;
            return (ms, true);
        }

        private static async ValueTask<(Stream Source, bool OwnsSource)> EnsureSeekableAsync(Stream stream, bool leaveOpen, CancellationToken ct)
        {
            if (stream.CanSeek)
            {
                return (stream, !leaveOpen);
            }
            MemoryStream ms = new();
            await stream.CopyToAsync(ms, ct).ConfigureAwait(false);
            if (!leaveOpen)
            {
                await stream.DisposeAsync().ConfigureAwait(false);
            }
            ms.Position = 0;
            return (ms, true);
        }

        internal static MemoryStream AsStream(ReadOnlyMemory<byte> data)
        {
            if (MemoryMarshal.TryGetArray(data, out ArraySegment<byte> segment))
            {
                return new MemoryStream(segment.Array!, segment.Offset, segment.Count, writable: false);
            }
            return new MemoryStream(data.ToArray(), writable: false);
        }

        [SkipLocalsInit]
        private static WorkbookStream BuildWorkbook(Stream source, bool ownsSource, ExcelReaderOptions options, ReadOnlyMemory<byte> memory = default)
        {
            if (source.Length < HeaderSize)
            {
                throw new InvalidDataException("The stream is not an OLE compound document.");
            }
            Span<byte> header = stackalloc byte[HeaderSize];
            ReadAt(source, 0, header);
            if (!header[..Signature.Length].SequenceEqual(Signature))
            {
                throw new InvalidDataException("The stream is not an OLE compound document.");
            }

            int sectorSize = 1 << ReadU16(header, 0x1E);
            int miniSectorSize = 1 << ReadU16(header, 0x20);
            int fatSectorCount = ReadI32(header, 0x2C);
            int firstDirectorySector = ReadI32(header, 0x30);
            int miniCutoff = ReadI32(header, 0x38);
            int firstMiniFatSector = ReadI32(header, 0x3C);
            int miniFatSectorCount = ReadI32(header, 0x40);
            int firstDifatSector = ReadI32(header, 0x44);
            int difatSectorCount = ReadI32(header, 0x48);

            if (sectorSize < HeaderSize || sectorSize > 4096 || miniSectorSize <= 0)
            {
                throw new InvalidDataException("Unsupported OLE sector size.");
            }
            // MS-CFB fixes the mini-stream cutoff at 4096 bytes. Without this bound a crafted header
            // could push miniCutoff toward int.MaxValue, letting the mini-stream branch below take a
            // multi-GB workbook.Size and materialize it as a single non-pooled byte[].
            if (miniCutoff != 4096)
            {
                throw new InvalidDataException("Unsupported OLE mini stream cutoff.");
            }
            // A file cannot hold more sectors than its length allows, so a FAT/DIFAT sector count above
            // that is a crafted header. Reject it before allocating, or `new int[fatSectorCount]` below
            // would let a bogus count force a multi-GB allocation / OOM on untrusted input.
            long maxSectors = source.Length / sectorSize;
            if (fatSectorCount < 0 || fatSectorCount > maxSectors ||
                difatSectorCount < 0 || difatSectorCount > maxSectors)
            {
                throw new InvalidDataException("The OLE FAT sector count is out of range.");
            }
            var fatSectorIds = new int[fatSectorCount];
            ReadDifat(source, header, sectorSize, fatSectorIds, firstDifatSector, difatSectorCount);
            int[] fatArray = ReadFat(source, sectorSize, fatSectorIds, out int fatLength);
            try
            {
                ReadOnlySpan<int> fat = fatArray.AsSpan(0, fatLength);
                byte[] directory = ReadChainBytes(source, sectorSize, fat, firstDirectorySector, -1);
                DirectoryEntry[] entries = ReadDirectory(directory);
                if (entries.Length == 0)
                {
                    throw new InvalidDataException("The OLE directory is empty.");
                }

                DirectoryEntry workbook = FindWorkbook(entries);

                // A stream cannot hold more content than the container's own byte length, so an
                // inflated Size field (the same attack class as fatSectorCount/difatSectorCount above)
                // is a crafted header — reject it before it drives an allocation or a chain walk sized
                // off it. The caller's byte budget applies here too, since this is the one choke point
                // both the mini-stream and chained/streamed branches below pass through.
                if (workbook.Size < 0 || workbook.Size > source.Length)
                {
                    throw new InvalidDataException("The OLE Workbook stream size exceeds the container.");
                }
                LimitChecks.ThrowIfEntryLengthExceeds(workbook.Size, options.MaxTotalDecompressedBytes, nameof(ExcelReaderOptions.MaxTotalDecompressedBytes));

                // Mini-stream workbooks (tiny, rare) are materialized; everything else streams.
                if (workbook.Size < miniCutoff && workbook.StartSector >= 0)
                {
                    int[] miniFat = firstMiniFatSector >= 0 && miniFatSectorCount > 0
                        ? ReadIntSectors(source, sectorSize, fat, firstMiniFatSector, miniFatSectorCount)
                        : [];
                    byte[] miniStream = entries[0].StartSector >= 0 && entries[0].Size > 0
                        ? ReadChainBytes(source, sectorSize, fat, entries[0].StartSector, (int)entries[0].Size)
                        : [];
                    if (miniFat.Length == 0 || miniStream.Length == 0)
                    {
                        throw new InvalidDataException("The OLE mini stream is missing.");
                    }
                    byte[] data = ReadMiniStream(miniStream, miniFat, miniSectorSize, workbook.StartSector, (int)workbook.Size);
                    if (ownsSource)
                    {
#pragma warning disable IDISP007 // Disposed only when this method owns the source; the in-memory workbook no longer needs it.
                        source.Dispose();
#pragma warning restore IDISP007
                    }
                    return WorkbookStream.InMemory(data);
                }

                int chainCount = SectorCount(workbook.Size, sectorSize);
                int[] chain = BuildChain(fat, workbook.StartSector, chainCount);
                if (!memory.IsEmpty)
                {
                    return WorkbookStream.Chained(memory, chain, chainCount, sectorSize, workbook.Size);
                }
                return WorkbookStream.Streamed(source, ownsSource, chain, chainCount, sectorSize, workbook.Size);
            }
            finally
            {
                ArrayPool<int>.Shared.Return(fatArray);
            }
        }

        private static DirectoryEntry FindWorkbook(ReadOnlySpan<DirectoryEntry> entries)
        {
            foreach (ref readonly var entry in entries)
            {
                if (entry.ObjectType == 2 &&
                    (entry.Name.Equals("Workbook", StringComparison.OrdinalIgnoreCase) ||
                     entry.Name.Equals("Book", StringComparison.OrdinalIgnoreCase)))
                {
                    return entry;
                }
            }
            throw new InvalidDataException("The OLE document does not contain a Workbook stream.");
        }

        private static int SectorCount(long size, int sectorSize)
        {
            return checked((int)((size + sectorSize - 1) / sectorSize));
        }

        [SuppressMessage("Performance", "HLQ013:Consider using 'foreach' loop instead of 'for' loop",
            Justification = "Not an iteration over fat; follows the sector linked-list, writing each hop into chain[i].")]
        // Rents the chain from the pool (oversized); WorkbookStream owns it for the read and returns it
        // in Dispose, bounding all access by the sectorCount it also receives (not chain.Length).
        private static int[] BuildChain(ReadOnlySpan<int> fat, int startSector, int sectorCount)
        {
            int[] chain = ArrayPool<int>.Shared.Rent(sectorCount);
            try
            {
                int sector = startSector;
                for (int i = 0; i < sectorCount; i++)
                {
                    if (sector is < 0 or EndOfChain)
                    {
                        throw new InvalidDataException("The OLE Workbook chain ended early.");
                    }
                    chain[i] = sector;
                    sector = NextSector(fat, sector);
                }
                return chain;
            }
            catch
            {
                ArrayPool<int>.Shared.Return(chain);
                throw;
            }
        }

        private static void ReadAt(Stream source, long offset, Span<byte> dest)
        {
            source.Seek(offset, SeekOrigin.Begin);
            source.ReadExactly(dest);
        }

        private static void ReadDifat(Stream source, ReadOnlySpan<byte> header, int sectorSize, Span<int> fatSectors, int firstDifatSector, int difatSectorCount)
        {
            int count = 0;
            for (int i = 0x4C; i < HeaderSize && count < fatSectors.Length; i += 4)
            {
                int sector = ReadI32(header, i);
                if (sector is >= 0 and not FreeSector)
                {
                    fatSectors[count++] = sector;
                }
            }

            int difat = firstDifatSector;
            byte[] difatSector = ArrayPool<byte>.Shared.Rent(sectorSize);
            try
            {
                for (int i = 0; i < difatSectorCount && difat >= 0 && count < fatSectors.Length; i++)
                {
                    ReadAt(source, SectorOffset(difat, sectorSize), difatSector);
                    int entries = (sectorSize / 4) - 1;
                    for (int j = 0; j < entries && count < fatSectors.Length; j++)
                    {
                        int sector = ReadI32(difatSector, j * 4);
                        if (sector is >= 0 and not FreeSector)
                        {
                            fatSectors[count++] = sector;
                        }
                    }
                    difat = ReadI32(difatSector, entries * 4);
                }

                if (count != fatSectors.Length)
                {
                    throw new InvalidDataException("The OLE DIFAT is incomplete.");
                }
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(difatSector);
            }
        }

        // Rents from the pool (returned in BuildWorkbook's finally); fatLength is the true entry count,
        // since the rented array is oversized — callers must bound reads by fatLength, not fat.Length.
        private static int[] ReadFat(Stream source, int sectorSize, ReadOnlySpan<int> fatSectorIds, out int fatLength)
        {
            int entriesPerSector = sectorSize / 4;
            fatLength = fatSectorIds.Length * entriesPerSector;
            int[] fat = ArrayPool<int>.Shared.Rent(fatLength);
            int index = 0;
            var sectorBuf = ArrayPool<byte>.Shared.Rent(sectorSize);
            try
            {
                foreach (ref readonly var sector in fatSectorIds)
                {
                    ReadAt(source, SectorOffset(sector, sectorSize), sectorBuf);
                    for (int i = 0; i < entriesPerSector; i++)
                    {
                        fat[index++] = ReadI32(sectorBuf, i * 4);
                    }
                }
                return fat;
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(sectorBuf);
            }
        }

        // Reads a FAT-chained stream into a byte[]. byteLimit < 0 means "until end of chain".
        private static byte[] ReadChainBytes(Stream source, int sectorSize, ReadOnlySpan<int> fat, int startSector, int byteLimit)
        {
            if (startSector < 0)
            {
                return [];
            }
            using MemoryStream ms = new();
            int sector = startSector;
            int written = 0;
            byte[] sectorBuf = ArrayPool<byte>.Shared.Rent(sectorSize);
            try
            {
                int sectorsRead = 0;
                while (sector is >= 0 and not EndOfChain && (byteLimit < 0 || written < byteLimit))
                {
                    if (sectorsRead++ >= fat.Length)
                    {
                        throw new InvalidDataException("OLE FAT chain contains a cycle.");
                    }
                    ReadAt(source, SectorOffset(sector, sectorSize), sectorBuf);
                    int take = byteLimit < 0 ? sectorSize : Math.Min(sectorSize, byteLimit - written);
                    ms.Write(sectorBuf, 0, take);
                    written += take;
                    sector = NextSector(fat, sector);
                }
                return ms.ToArray();
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(sectorBuf);
            }
        }

        private static int[] ReadIntSectors(Stream source, int sectorSize, ReadOnlySpan<int> fat, int firstSector, int sectorCount)
        {
            byte[] data = ReadChainBytes(source, sectorSize, fat, firstSector, checked(sectorCount * sectorSize));
            int[] result = new int[data.Length / 4];
            for (int i = 0; i < result.Length; i++)
            {
                result[i] = ReadI32(data, i * 4);
            }
            return result;
        }

        private static byte[] ReadMiniStream(ReadOnlySpan<byte> miniStream, ReadOnlySpan<int> miniFat, int miniSectorSize, int startSector, int size)
        {
            if (size < 0 || size > miniStream.Length)
            {
                throw new InvalidDataException("Invalid OLE mini stream size.");
            }
            byte[] result = new byte[size];
            int sector = startSector;
            int written = 0;
            while (sector is >= 0 and not EndOfChain && written < result.Length)
            {
                int offset = checked(sector * miniSectorSize);
                if ((uint)offset >= (uint)miniStream.Length)
                {
                    throw new InvalidDataException("Invalid OLE mini sector chain.");
                }
                int take = Math.Min(miniSectorSize, result.Length - written);
                miniStream.Slice(offset, take).CopyTo(result.AsSpan(written));
                written += take;
                if ((uint)sector >= (uint)miniFat.Length)
                {
                    throw new InvalidDataException("Invalid OLE mini FAT chain.");
                }
                sector = miniFat[sector];
            }
            return result;
        }

        private static int NextSector(ReadOnlySpan<int> fat, int sector)
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

        private static long SectorOffset(int sector, int sectorSize)
        {
            if (sector < 0)
            {
                throw new InvalidDataException("Invalid OLE sector offset.");
            }
            return ((long)sector + 1) * sectorSize;
        }

        private readonly record struct DirectoryEntry(string Name, byte ObjectType, int StartSector, long Size);
    }
}
