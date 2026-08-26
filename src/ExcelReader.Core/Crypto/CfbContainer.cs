using System.Buffers;
using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using ExcelReader.Core.Reader;
using static ExcelReader.Core.Reader.Biff12;

namespace ExcelReader.Core.Crypto
{
    // Parses the OLE/CFB container metadata (header, FAT, directory, mini-FAT/stream) by seeking a
    // seekable source, never materializing the whole file, and then exposes the directory so callers
    // can look up any named stream by name (e.g. "Workbook"/"Book" for XLS, "EncryptionInfo"/
    // "EncryptedPackage" for an encrypted OOXML package). Extracted verbatim from
    // XlsCompoundFile.BuildWorkbook, which used to hardwire the directory lookup to a single stream.
    [ExcludeFromCodeCoverage(Justification = "Covered through XlsReader integration tests; most uncovered paths are corrupt-OLE guard rails.")]
    internal sealed class CfbContainer : IDisposable
    {
        private const int HeaderSize = 512;
        private const int EndOfChain = unchecked((int)0xFFFFFFFE);
        private const int FatSector = unchecked((int)0xFFFFFFFD);
        private const int FreeSector = unchecked((int)0xFFFFFFFF);

        // Internal (not private) so XlsCompoundFile.BuildWorkbook — the sole consumer that still needs
        // to reimplement the mini-cutoff / Chained / Streamed selection over the raw sector data — can
        // read them directly, exactly as it did when this state lived inline in that method.
        internal readonly Stream Source;
        internal readonly bool OwnsSource;
        internal readonly ReadOnlyMemory<byte> Memory;
        internal readonly ExcelReaderOptions Options;
        internal readonly int SectorSize;
        internal readonly int MiniSectorSize;
        internal readonly int MiniCutoff;
        internal readonly int[] FatArray;
        internal readonly int FatLength;
        internal readonly DirectoryEntry[] Entries;
        internal readonly int FirstMiniFatSector;
        internal readonly int MiniFatSectorCount;
        private bool _disposed;
        private bool _fatReturned;

        private CfbContainer(
            Stream source,
            bool ownsSource,
            ReadOnlyMemory<byte> memory,
            ExcelReaderOptions options,
            int sectorSize,
            int miniSectorSize,
            int miniCutoff,
            int[] fatArray,
            int fatLength,
            DirectoryEntry[] entries,
            int firstMiniFatSector,
            int miniFatSectorCount)
        {
            Source = source;
            OwnsSource = ownsSource;
            Memory = memory;
            Options = options;
            SectorSize = sectorSize;
            MiniSectorSize = miniSectorSize;
            MiniCutoff = miniCutoff;
            FatArray = fatArray;
            FatLength = fatLength;
            Entries = entries;
            FirstMiniFatSector = firstMiniFatSector;
            MiniFatSectorCount = miniFatSectorCount;
        }

        internal ReadOnlySpan<int> Fat => FatArray.AsSpan(0, FatLength);

        [SkipLocalsInit]
        internal static CfbContainer Parse(Stream source, bool ownsSource, ExcelReaderOptions options, ReadOnlyMemory<byte> memory = default)
        {
            if (source.Length < HeaderSize)
            {
                throw new InvalidDataException("The stream is not an OLE compound document.");
            }
            Span<byte> header = stackalloc byte[HeaderSize];
            ReadAt(source, 0, header);
            if (!header[..XlsCompoundFile.Signature.Length].SequenceEqual(XlsCompoundFile.Signature))
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

            // [MS-CFB] fixes the mini sector size at 64 bytes (shift = 6); nothing bounded the upper end
            // here before, and an unchecked shift amount from the file could still land the result well
            // above 64 (int shift amounts wrap mod 32, so this never overflows, but a crafted header
            // could pick any power-of-two result and later drive checked(sector * miniSectorSize) into an
            // avoidable OverflowException instead of a graceful rejection at the source).
            if (sectorSize < HeaderSize || sectorSize > 4096 || miniSectorSize != 64)
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
            // A file cannot hold more sectors than its length allows, so a FAT/DIFAT/mini-FAT sector
            // count above that is a crafted header. Reject it before allocating, or `new
            // int[fatSectorCount]` below would let a bogus count force a multi-GB allocation / OOM on
            // untrusted input.
            //
            // miniFatSectorCount belongs here for a second reason: it is multiplied by sectorSize in
            // ReadIntSectors, and a large value overflows that `checked` product into an
            // OverflowException — a leaked arithmetic fault rather than the InvalidDataException a
            // malformed file is supposed to produce. Found by the XLS fuzz target.
            long maxSectors = source.Length / sectorSize;
            if (fatSectorCount < 0 || fatSectorCount > maxSectors ||
                difatSectorCount < 0 || difatSectorCount > maxSectors ||
                miniFatSectorCount < 0 || miniFatSectorCount > maxSectors)
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

                return new CfbContainer(
                    source,
                    ownsSource,
                    memory,
                    options,
                    sectorSize,
                    miniSectorSize,
                    miniCutoff,
                    fatArray,
                    fatLength,
                    entries,
                    firstMiniFatSector,
                    miniFatSectorCount);
            }
            catch
            {
                ArrayPool<int>.Shared.Return(fatArray);
                throw;
            }
        }

        internal bool ContainsStream(string name)
        {
            return TryFindEntry(name, out _);
        }

        internal long StreamLength(string name)
        {
            if (!TryFindEntry(name, out DirectoryEntry entry))
            {
                throw new InvalidDataException($"The OLE document does not contain a '{name}' stream.");
            }
            return entry.Size;
        }

        internal byte[] ReadStream(string name, long maxBytes)
        {
            if (!TryFindEntry(name, out DirectoryEntry entry))
            {
                throw new InvalidDataException($"The OLE document does not contain a '{name}' stream.");
            }

            // A stream cannot hold more content than the container's own byte length, so an inflated
            // Size field (the same attack class as fatSectorCount/difatSectorCount above) is a crafted
            // header — reject it before it drives an allocation or a chain walk sized off it.
            if (entry.Size < 0 || entry.Size > Source.Length)
            {
                throw new InvalidDataException($"The OLE '{name}' stream size exceeds the container.");
            }
            LimitChecks.ThrowIfEntryLengthExceeds(entry.Size, maxBytes, nameof(maxBytes));

            ReadOnlySpan<int> fatSpan = Fat;
            if (entry.Size < MiniCutoff && entry.StartSector >= 0)
            {
                byte[] miniStream = ReadMiniStreamData(fatSpan, out int[] miniFat);
                return ReadMiniStream(miniStream, miniFat, MiniSectorSize, entry.StartSector, (int)entry.Size);
            }

            return ReadChainBytes(Source, SectorSize, fatSpan, entry.StartSector, (int)entry.Size);
        }

        internal Stream OpenStreamView(string name)
        {
            if (!TryFindEntry(name, out DirectoryEntry entry))
            {
                throw new InvalidDataException($"The OLE document does not contain a '{name}' stream.");
            }
            if (entry.Size < 0 || entry.Size > Source.Length)
            {
                throw new InvalidDataException($"The OLE '{name}' stream size exceeds the container.");
            }

            ReadOnlySpan<int> fatSpan = Fat;
            if (entry.Size < MiniCutoff && entry.StartSector >= 0)
            {
                byte[] miniStream = ReadMiniStreamData(fatSpan, out int[] miniFat);
                byte[] data = ReadMiniStream(miniStream, miniFat, MiniSectorSize, entry.StartSector, (int)entry.Size);
                return new MemoryStream(data, writable: false);
            }

            int chainCount = SectorCount(entry.Size, SectorSize);
            int[] chain = BuildChain(fatSpan, entry.StartSector, chainCount);
            return new CfbStreamView(Source, chain, SectorSize, entry.Size);
        }

        // Shared by ReadStream and OpenStreamView's mini-stream branches: rebuilds the mini-FAT (from
        // the FAT-chained mini-FAT sectors) and the root entry's materialized mini-stream, the same way
        // BuildWorkbook used to do it inline.
        private byte[] ReadMiniStreamData(ReadOnlySpan<int> fatSpan, out int[] miniFat)
        {
            miniFat = FirstMiniFatSector >= 0 && MiniFatSectorCount > 0
                ? ReadIntSectors(Source, SectorSize, fatSpan, FirstMiniFatSector, MiniFatSectorCount)
                : [];
            // Entries[0].Size (the root storage entry's mini-stream length) is a long; a value above
            // int.MaxValue would truncate through the (int) cast into a negative byteLimit, which
            // ReadChainBytes interprets as "read the entire chain" instead of "read N bytes" — bounded
            // safely by the cycle check below, but a silent semantic flip worth closing.
            if (Entries[0].Size > int.MaxValue)
            {
                throw new InvalidDataException("The OLE root entry size exceeds the container.");
            }
            byte[] miniStream = Entries[0].StartSector >= 0 && Entries[0].Size > 0
                ? ReadChainBytes(Source, SectorSize, fatSpan, Entries[0].StartSector, (int)Entries[0].Size)
                : [];
            if (miniFat.Length == 0 || miniStream.Length == 0)
            {
                throw new InvalidDataException("The OLE mini stream is missing.");
            }
            return miniStream;
        }

        internal bool TryFindEntry(string name, out DirectoryEntry found)
        {
            foreach (ref readonly var entry in Entries.AsSpan())
            {
                if (entry.ObjectType == 2 && entry.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    found = entry;
                    return true;
                }
            }
            found = default;
            return false;
        }

        // Returns the pooled FAT array without touching `Source`. XlsCompoundFile.BuildWorkbook calls
        // this instead of Dispose(): it hands `Source` (and its OwnsSource flag) off to the
        // WorkbookStream it returns (or disposes it itself, for the mini-stream branch), so the
        // container must not also dispose it here.
        internal void ReturnFatBuffer()
        {
            if (_fatReturned)
            {
                return;
            }
            _fatReturned = true;
            ArrayPool<int>.Shared.Return(FatArray);
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            ReturnFatBuffer();
            if (OwnsSource)
            {
#pragma warning disable IDISP007
                Source.Dispose();
#pragma warning restore IDISP007
            }
        }

        internal static int SectorCount(long size, int sectorSize)
        {
            return checked((int)((size + sectorSize - 1) / sectorSize));
        }

        [SuppressMessage("Performance", "HLQ013:Consider using 'foreach' loop instead of 'for' loop",
            Justification = "Not an iteration over fat; follows the sector linked-list, writing each hop into chain[i].")]
        // Rents the chain from the pool (oversized); WorkbookStream/CfbStreamView owns it for the read
        // and returns it in Dispose, bounding all access by the sectorCount it also receives (not
        // chain.Length).
        internal static int[] BuildChain(ReadOnlySpan<int> fat, int startSector, int sectorCount)
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

        // Every sector-based read in this file (FAT, DIFAT, chain walks) funnels through here with an
        // offset derived from a sector id read straight from the file. SectorOffset already rejects a
        // negative sector, but a huge positive one (still a valid int, e.g. from a single flipped byte)
        // passed that check and reached Stream.Seek/ReadExactly directly — surfacing as a raw
        // ArgumentOutOfRangeException or EndOfStreamException instead of the graceful InvalidDataException
        // every other bound in this file already throws. This is the one choke point all three callers
        // share, so the bound belongs here rather than duplicated at each call site.
        private static void ReadAt(Stream source, long offset, Span<byte> dest)
        {
            if (offset < 0 || offset > source.Length - dest.Length)
            {
                throw new InvalidDataException("The OLE compound file references a sector outside the container.");
            }
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

        // Rents from the pool (returned in the container's Dispose); fatLength is the true entry count,
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
        internal static byte[] ReadChainBytes(Stream source, int sectorSize, ReadOnlySpan<int> fat, int startSector, int byteLimit)
        {
            if (startSector < 0)
            {
                return [];
            }
            using MemoryStream ms = new();
            int sector = startSector;
            int written = 0;
            byte[] sectorBuf = ArrayPool<byte>.Shared.Rent(sectorSize);
            // Counting iterations (as this used to) lets a 2-sector cycle (fat[a]=b, fat[b]=a) run the
            // full fat.Length before tripping, writing sectorSize bytes per hop into this unbounded
            // MemoryStream — up to ~1000x amplification on a file whose FAT happens to have many
            // entries. Tracking visited sectors instead catches a cycle after at most fat.Length
            // distinct sectors, which is the true worst case for an acyclic chain too.
            bool[] visited = ArrayPool<bool>.Shared.Rent(Math.Max(1, fat.Length));
            Array.Clear(visited, 0, fat.Length);
            try
            {
                while (sector is >= 0 and not EndOfChain && (byteLimit < 0 || written < byteLimit))
                {
                    if ((uint)sector >= (uint)fat.Length || visited[sector])
                    {
                        throw new InvalidDataException("OLE FAT chain contains a cycle.");
                    }
                    visited[sector] = true;
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
                ArrayPool<bool>.Shared.Return(visited);
            }
        }

        internal static int[] ReadIntSectors(Stream source, int sectorSize, ReadOnlySpan<int> fat, int firstSector, int sectorCount)
        {
            byte[] data = ReadChainBytes(source, sectorSize, fat, firstSector, checked(sectorCount * sectorSize));
            int[] result = new int[data.Length / 4];
            for (int i = 0; i < result.Length; i++)
            {
                result[i] = ReadI32(data, i * 4);
            }
            return result;
        }

        internal static byte[] ReadMiniStream(ReadOnlySpan<byte> miniStream, ReadOnlySpan<int> miniFat, int miniSectorSize, int startSector, int size)
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
                // [MS-CFB] caps a directory entry's name length at 64 bytes (including the null
                // terminator); nothing enforced that here, so a crafted value up to 65535 either threw
                // a raw ArgumentOutOfRangeException slicing this fixed 128-byte slot, or (for values
                // between 66 and 130) silently read adjacent slot fields into the name string.
                if (nameBytes < 0 || nameBytes > 64)
                {
                    throw new InvalidDataException("The OLE directory entry name length is out of range.");
                }
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

        internal readonly record struct DirectoryEntry(string Name, byte ObjectType, int StartSector, long Size);

        // Read-only, seekable view over a FAT-chained CFB stream. This is what lets ZipArchive seek to
        // the central directory of an EncryptedPackage stream without the caller materializing the
        // whole (potentially large) decrypted package up front.
        private sealed class CfbStreamView : Stream
        {
            // Borrowed from the owning CfbContainer, which outlives this view and owns disposal — do
            // not dispose here.
            [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed", Justification = "Borrowed, not owned.")]
            private readonly Stream _source;
            private readonly int[] _chain;
            private readonly int _sectorSize;
            private long _position;
            private bool _disposed;

            internal CfbStreamView(Stream source, int[] chain, int sectorSize, long length)
            {
                _source = source;
                _chain = chain;
                _sectorSize = sectorSize;
                Length = length;
            }

            public override bool CanRead => true;

            public override bool CanSeek => true;

            public override bool CanWrite => false;

            public override long Length { get; }

            public override long Position
            {
                get => _position;
                set
                {
                    ArgumentOutOfRangeException.ThrowIfNegative(value);
                    _position = value;
                }
            }

            public override int Read(byte[] buffer, int offset, int count)
            {
                return Read(buffer.AsSpan(offset, count));
            }

            public override int Read(Span<byte> buffer)
            {
                if (_position >= Length)
                {
                    return 0;
                }
                long remaining = Length - _position;
                int toRead = (int)Math.Min(buffer.Length, remaining);
                int totalRead = 0;
                while (totalRead < toRead)
                {
                    long absolute = _position + totalRead;
                    int sectorIndex = (int)(absolute / _sectorSize);
                    int sectorOffset = (int)(absolute % _sectorSize);
                    int sector = _chain[sectorIndex];
                    long fileOffset = ((long)sector + 1) * _sectorSize + sectorOffset;
                    int chunk = Math.Min(toRead - totalRead, _sectorSize - sectorOffset);
                    _source.Seek(fileOffset, SeekOrigin.Begin);
                    _source.ReadExactly(buffer.Slice(totalRead, chunk));
                    totalRead += chunk;
                }
                _position += totalRead;
                return totalRead;
            }

            public override long Seek(long offset, SeekOrigin origin)
            {
                long newPosition = origin switch
                {
                    SeekOrigin.Begin => offset,
                    SeekOrigin.Current => _position + offset,
                    SeekOrigin.End => Length + offset,
                    _ => throw new ArgumentOutOfRangeException(nameof(origin)),
                };
                if (newPosition < 0)
                {
                    throw new IOException("An attempt was made to move the position before the beginning of the stream.");
                }
                _position = newPosition;
                return _position;
            }

            public override void Flush()
            {
            }

            public override void SetLength(long value)
            {
                throw new NotSupportedException();
            }

            public override void Write(byte[] buffer, int offset, int count)
            {
                throw new NotSupportedException();
            }

            protected override void Dispose(bool disposing)
            {
                if (!_disposed && disposing)
                {
                    _disposed = true;
                    ArrayPool<int>.Shared.Return(_chain);
                }
                base.Dispose(disposing);
            }
        }
    }
}
