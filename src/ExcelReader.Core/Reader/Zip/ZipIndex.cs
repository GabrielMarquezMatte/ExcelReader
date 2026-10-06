using System.Buffers;
using System.Buffers.Binary;
using System.IO.Compression;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Sources;

namespace ExcelReader.Core.Reader.Zip
{
    [StructLayout(LayoutKind.Auto)]
    internal readonly struct ZipEntryRef
    {
        internal int NameStart { get; init; }
        internal int NameLength { get; init; }
        internal long LocalHeaderOffset { get; init; }
        internal long CompressedSize { get; init; }
        internal long UncompressedSize { get; init; }
        internal ushort Method { get; init; }
        internal ushort Flags { get; init; }
    }

    [StructLayout(LayoutKind.Auto)]
    internal readonly struct ZipPart : IDisposable
    {
        private readonly byte[]? _rented;

        internal ZipPart(ReadOnlyMemory<byte> memory, byte[]? rented)
        {
            Memory = memory;
            _rented = rented;
        }

        internal ReadOnlyMemory<byte> Memory { get; }

        public void Dispose()
        {
            if (_rented is not null)
            {
                ArrayPool<byte>.Shared.Return(_rented);
            }
        }
    }

    internal sealed class ZipIndex : IDisposable
    {
        private const int EocdFixedSize = 22;
        private const int Zip64LocatorSize = 20;
        private const int Zip64EocdFixedSize = 56;
        private const int CentralDirectoryFixedSize = 46;
        private const int LocalHeaderFixedSize = 30;
        private const int Zip64EocdLocatorSignature = 0x07064b50;
        private const int Zip64EocdSignature = 0x06064b50;
        private const int CentralDirectorySignature = 0x02014b50;
        private const int LocalFileHeaderSignature = 0x04034b50;
        private const ushort EncryptedFlag = 0x0001;
        private const ushort Zip64ExtraFieldId = 0x0001;
        private const uint Zip64SentinelU32 = 0xFFFFFFFFu;

        private static ReadOnlySpan<byte> EocdSignatureBytes => [0x50, 0x4B, 0x05, 0x06];

        private const int TailWindowSize = 65535 + EocdFixedSize + Zip64LocatorSize;

        private readonly ByteSource _source;
        private readonly ReadOnlyMemory<byte> _file;
        private readonly ReadOnlyMemory<byte> _directory;
        private readonly byte[]? _directoryRented;
        private readonly ZipEntryRef[] _entries;
        private int _disposed;

        private ZipIndex(ByteSource source, ReadOnlyMemory<byte> directory, byte[]? directoryRented, ZipEntryRef[] entries, int count)
        {
            _source = source;
            HasMemory = source.TryGetMemory(out _file);
            _directory = directory;
            _directoryRented = directoryRented;
            _entries = entries;
            Count = count;
        }

        internal int Count { get; }

        internal bool HasMemory { get; }

        internal static ZipIndex Create(ReadOnlyMemory<byte> file, ExcelReaderOptions options)
        {
            return Create(ByteSource.FromMemory(file), options);
        }

        internal static ZipIndex Create(ByteSource source, ExcelReaderOptions options)
        {
            byte[]? directoryRented = null;
            try
            {
                long length = source.Length;
                int tailSize = (int)Math.Min(length, TailWindowSize);
                long cdOffset;
                long cdSize;
                long declaredCount;
                long zip64EocdOffset;
                if (source.TryGetMemory(out ReadOnlyMemory<byte> file))
                {
                    (cdOffset, cdSize, declaredCount, zip64EocdOffset) = ParseTail(file.Span[(file.Length - tailSize)..]);
                }
                else
                {
                    byte[] tail = ArrayPool<byte>.Shared.Rent(Math.Max(1, tailSize));
                    try
                    {
                        source.ReadExactly(length - tailSize, tail.AsSpan(0, tailSize));
                        (cdOffset, cdSize, declaredCount, zip64EocdOffset) = ParseTail(tail.AsSpan(0, tailSize));
                    }
                    finally
                    {
                        ArrayPool<byte>.Shared.Return(tail);
                    }
                }
                if (zip64EocdOffset >= 0)
                {
                    ThrowIfZip64RecordOutOfRange(zip64EocdOffset, length);
                    Span<byte> record = stackalloc byte[Zip64EocdFixedSize];
                    source.ReadExactly(zip64EocdOffset, record);
                    (cdOffset, cdSize, declaredCount) = ReadZip64Eocd(record, 0);
                }
                ThrowIfDirectoryOutOfRange(cdOffset, cdSize, length);

                ReadOnlyMemory<byte> directory;
                if (source.TryGetMemory(out file))
                {
                    directory = file.Slice((int)cdOffset, (int)cdSize);
                }
                else
                {
                    directoryRented = ArrayPool<byte>.Shared.Rent(Math.Max(1, (int)cdSize));
                    source.ReadExactly(cdOffset, directoryRented.AsSpan(0, (int)cdSize));
                    directory = directoryRented.AsMemory(0, (int)cdSize);
                }
                return Build(source, directory, directoryRented, declaredCount, options);
            }
            catch
            {
                if (directoryRented is not null)
                {
                    ArrayPool<byte>.Shared.Return(directoryRented);
                }
                source.Dispose();
                throw;
            }
        }

        internal static async ValueTask<ZipIndex> CreateAsync(ByteSource source, ExcelReaderOptions options, CancellationToken ct)
        {
            if (source.TryGetMemory(out _))
            {
                return Create(source, options);
            }
            byte[]? directoryRented = null;
            try
            {
                long length = source.Length;
                int tailSize = (int)Math.Min(length, TailWindowSize);
                long cdOffset;
                long cdSize;
                long declaredCount;
                long zip64EocdOffset;
                byte[] tail = ArrayPool<byte>.Shared.Rent(Math.Max(1, tailSize));
                try
                {
                    await source.ReadExactlyAsync(length - tailSize, tail.AsMemory(0, tailSize), ct).ConfigureAwait(false);
                    (cdOffset, cdSize, declaredCount, zip64EocdOffset) = ParseTail(tail.AsSpan(0, tailSize));
                }
                finally
                {
                    ArrayPool<byte>.Shared.Return(tail);
                }
                if (zip64EocdOffset >= 0)
                {
                    ThrowIfZip64RecordOutOfRange(zip64EocdOffset, length);
                    byte[] record = ArrayPool<byte>.Shared.Rent(Zip64EocdFixedSize);
                    try
                    {
                        await source.ReadExactlyAsync(zip64EocdOffset, record.AsMemory(0, Zip64EocdFixedSize), ct).ConfigureAwait(false);
                        (cdOffset, cdSize, declaredCount) = ReadZip64Eocd(record.AsSpan(0, Zip64EocdFixedSize), 0);
                    }
                    finally
                    {
                        ArrayPool<byte>.Shared.Return(record);
                    }
                }
                ThrowIfDirectoryOutOfRange(cdOffset, cdSize, length);
                directoryRented = ArrayPool<byte>.Shared.Rent(Math.Max(1, (int)cdSize));
                await source.ReadExactlyAsync(cdOffset, directoryRented.AsMemory(0, (int)cdSize), ct).ConfigureAwait(false);
                return Build(source, directoryRented.AsMemory(0, (int)cdSize), directoryRented, declaredCount, options);
            }
            catch
            {
                if (directoryRented is not null)
                {
                    ArrayPool<byte>.Shared.Return(directoryRented);
                }
                source.Dispose();
                throw;
            }
        }

        private static (long CdOffset, long CdSize, long Count, long Zip64EocdOffset) ParseTail(ReadOnlySpan<byte> tail)
        {
            long eocdOffset = FindEocd(tail);
            (long cdOffset, long cdSize, long declaredCount) = ReadEocdRecord(tail, eocdOffset, out long zip64EocdOffset);
            return (cdOffset, cdSize, declaredCount, zip64EocdOffset);
        }

        private static void ThrowIfZip64RecordOutOfRange(long offset, long length)
        {
            if (offset < 0 || offset > length - Zip64EocdFixedSize)
            {
                throw new InvalidDataException("The ZIP64 end of central directory record is out of range.");
            }
        }

        private static void ThrowIfDirectoryOutOfRange(long cdOffset, long cdSize, long length)
        {
            if (cdOffset < 0 || cdSize < 0 || cdOffset > length - cdSize)
            {
                throw new InvalidDataException("The ZIP central directory is out of range.");
            }
            if (cdSize > Array.MaxLength)
            {
                throw new ExcelLimitExceededException("ArrayMaxLength", Array.MaxLength, cdSize);
            }
        }

        private static ZipIndex Build(ByteSource source, ReadOnlyMemory<byte> directory, byte[]? directoryRented, long declaredCount, ExcelReaderOptions options)
        {
            long maxHint = Math.Max(16, options.MaxZipEntries > 0 ? options.MaxZipEntries : 65_536);
            ZipEntryRef[] entries = ArrayPool<ZipEntryRef>.Shared.Rent((int)Math.Clamp(declaredCount, 16, maxHint));
            try
            {
                int count = WalkCentralDirectory(directory.Span, ref entries, options);
                ThrowIfDuplicateEntryNames(directory, entries, count);
                return new ZipIndex(source, directory, directoryRented, entries, count);
            }
            catch
            {
                ArrayPool<ZipEntryRef>.Shared.Return(entries);
                throw;
            }
        }

        private static void ThrowIfDuplicateEntryNames(ReadOnlyMemory<byte> directory, ZipEntryRef[] entries, int count)
        {
            if (count <= 1)
            {
                return;
            }
            int[] order = ArrayPool<int>.Shared.Rent(count);
            try
            {
                for (int i = 0; i < count; i++)
                {
                    order[i] = i;
                }
                Array.Sort(order, 0, count, Comparer<int>.Create((a, b) =>
                {
                    ReadOnlySpan<byte> span = directory.Span;
                    ref readonly ZipEntryRef ea = ref entries[a];
                    ref readonly ZipEntryRef eb = ref entries[b];
                    return span.Slice(ea.NameStart, ea.NameLength).SequenceCompareTo(span.Slice(eb.NameStart, eb.NameLength));
                }));
                ReadOnlySpan<byte> directorySpan = directory.Span;
                for (int i = 1; i < count; i++)
                {
                    ref readonly ZipEntryRef prev = ref entries[order[i - 1]];
                    ref readonly ZipEntryRef curr = ref entries[order[i]];
                    if (prev.NameLength == curr.NameLength &&
                        directorySpan.Slice(prev.NameStart, prev.NameLength).SequenceEqual(directorySpan.Slice(curr.NameStart, curr.NameLength)))
                    {
                        throw new InvalidDataException("The ZIP central directory contains a duplicate entry name.");
                    }
                }
            }
            finally
            {
                ArrayPool<int>.Shared.Return(order);
            }
        }

        internal bool TryGetEntry(ReadOnlySpan<byte> utf8Name, out ZipEntryRef entry)
        {
            ReadOnlySpan<byte> directory = _directory.Span;
            foreach (ref readonly ZipEntryRef candidate in _entries.AsSpan(0, Count))
            {
                if (directory.Slice(candidate.NameStart, candidate.NameLength).SequenceEqual(utf8Name))
                {
                    entry = candidate;
                    return true;
                }
            }
            entry = default;
            return false;
        }

        internal ZipPart OpenPart(in ZipEntryRef entry, DecompressedByteCounter counter, string entryLimitName = "", long entryLimit = 0)
        {
            ThrowIfPartTooLarge(entry, counter, entryLimitName, entryLimit);
            long dataOffset = ResolveDataOffset(entry);
            ThrowIfDataOutOfRange(dataOffset, entry.CompressedSize);
            ZipPart part;
            if (HasMemory)
            {
                ReadOnlyMemory<byte> compressed = CompressedSlice(dataOffset, entry.CompressedSize);
                part = entry.Method switch
                {
                    0 => new ZipPart(compressed, rented: null),
                    8 => InflateToPart(compressed, entry.UncompressedSize),
                    _ => throw new NotSupportedException($"Unsupported ZIP compression method: {entry.Method}."),
                };
            }
            else
            {
                ThrowIfMethodUnsupported(entry.Method);
                part = ReadPart(dataOffset, entry);
            }
            counter.Add(entry.UncompressedSize);
            return part;
        }

        internal async ValueTask<ZipPart> OpenPartAsync(
            ZipEntryRef entry, DecompressedByteCounter counter, CancellationToken ct, string entryLimitName = "", long entryLimit = 0)
        {
            if (HasMemory)
            {
                return OpenPart(entry, counter, entryLimitName, entryLimit);
            }
            ThrowIfPartTooLarge(entry, counter, entryLimitName, entryLimit);
            long dataOffset = await ResolveDataOffsetAsync(entry, ct).ConfigureAwait(false);
            ThrowIfDataOutOfRange(dataOffset, entry.CompressedSize);
            ThrowIfMethodUnsupported(entry.Method);
            int size = PartLength(entry);
            byte[] rented = ArrayPool<byte>.Shared.Rent(Math.Max(1, size));
            try
            {
                ByteSourceStream raw = new(_source, dataOffset, entry.CompressedSize);
                Stream content = entry.Method == 0 ? raw : new DeflateStream(raw, CompressionMode.Decompress);
                await using (content.ConfigureAwait(false))
                {
                    try
                    {
                        await content.ReadExactlyAsync(rented.AsMemory(0, size), ct).ConfigureAwait(false);
                    }
                    catch (EndOfStreamException ex)
                    {
                        throw new InvalidDataException("The ZIP entry produced less data than its declared uncompressed size.", ex);
                    }
                    if (await content.ReadAsync(new byte[1], ct).ConfigureAwait(false) != 0)
                    {
                        throw new InvalidDataException("The ZIP entry produced more data than its declared uncompressed size.");
                    }
                }
            }
            catch
            {
                ArrayPool<byte>.Shared.Return(rented);
                throw;
            }
            counter.Add(entry.UncompressedSize);
            return new ZipPart(rented.AsMemory(0, size), rented);
        }

        internal ValueTask<ZipPart> OpenPartOrDefaultAsync(ReadOnlySpan<byte> utf8Name, DecompressedByteCounter counter, CancellationToken ct)
        {
            return TryGetEntry(utf8Name, out ZipEntryRef entry) ? OpenPartAsync(entry, counter, ct) : new ValueTask<ZipPart>(default(ZipPart));
        }

        internal ZipPart OpenPartOrDefault(ReadOnlySpan<byte> utf8Name, DecompressedByteCounter counter)
        {
            return TryGetEntry(utf8Name, out ZipEntryRef entry) ? OpenPart(entry, counter) : default;
        }

        internal LimitedReadStream OpenEntryStream(
            in ZipEntryRef entry, DecompressedByteCounter counter, ExcelReaderOptions options, string entryLimitName = "", long entryLimit = 0)
        {
            long dataOffset = ResolveDataOffset(entry);
            return WrapEntryStream(entry, dataOffset, counter, options, entryLimitName, entryLimit);
        }

        internal async ValueTask<LimitedReadStream> OpenEntryStreamAsync(
            ZipEntryRef entry, DecompressedByteCounter counter, ExcelReaderOptions options, CancellationToken ct,
            string entryLimitName = "", long entryLimit = 0)
        {
            long dataOffset = HasMemory ? ResolveDataOffset(entry) : await ResolveDataOffsetAsync(entry, ct).ConfigureAwait(false);
            return WrapEntryStream(entry, dataOffset, counter, options, entryLimitName, entryLimit);
        }

        private LimitedReadStream WrapEntryStream(
            in ZipEntryRef entry, long dataOffset, DecompressedByteCounter counter, ExcelReaderOptions options, string entryLimitName, long entryLimit)
        {
            ThrowIfDataOutOfRange(dataOffset, entry.CompressedSize);
            ThrowIfMethodUnsupported(entry.Method);
            Stream raw = HasMemory
                ? ToReadableMemoryStream(CompressedSlice(dataOffset, entry.CompressedSize))
                : new ByteSourceStream(_source, dataOffset, entry.CompressedSize);
            Stream opened = entry.Method == 0 ? raw : new DeflateStream(raw, CompressionMode.Decompress);
            return WorkbookLookups.Wrap(opened, counter, options, entryLimitName, entryLimit, entry.UncompressedSize);
        }

        private static void ThrowIfPartTooLarge(in ZipEntryRef entry, DecompressedByteCounter counter, string entryLimitName, long entryLimit)
        {
            LimitChecks.ThrowIfEntryLengthExceeds(entry.UncompressedSize, counter.Remaining, nameof(ExcelReaderOptions.MaxTotalDecompressedBytes));
            if (entryLimit > 0)
            {
                LimitChecks.ThrowIfEntryLengthExceeds(entry.UncompressedSize, entryLimit, entryLimitName);
            }
            if (entry.UncompressedSize > Array.MaxLength)
            {
                throw new ExcelLimitExceededException("ArrayMaxLength", Array.MaxLength, entry.UncompressedSize);
            }
        }

        private static void ThrowIfMethodUnsupported(ushort method)
        {
            if (method is not (0 or 8))
            {
                throw new NotSupportedException($"Unsupported ZIP compression method: {method}.");
            }
        }

        private void ThrowIfDataOutOfRange(long dataOffset, long compressedSize)
        {
            long length = _source.Length;
            if (dataOffset < 0 || dataOffset > length || compressedSize < 0 || compressedSize > length - dataOffset)
            {
                throw new InvalidDataException("The ZIP entry data runs past the end of the file.");
            }
        }

        private ReadOnlyMemory<byte> CompressedSlice(long dataOffset, long compressedSize)
        {
            if (compressedSize > Array.MaxLength)
            {
                throw new ExcelLimitExceededException("ArrayMaxLength", Array.MaxLength, compressedSize);
            }
            return _file.Slice((int)dataOffset, (int)compressedSize);
        }

        private static int PartLength(in ZipEntryRef entry)
        {
            long length = entry.Method == 0 ? entry.CompressedSize : entry.UncompressedSize;
            if (length > Array.MaxLength)
            {
                throw new ExcelLimitExceededException("ArrayMaxLength", Array.MaxLength, length);
            }
            return (int)length;
        }

        private ZipPart ReadPart(long dataOffset, in ZipEntryRef entry)
        {
            int size = PartLength(entry);
            byte[] rented = ArrayPool<byte>.Shared.Rent(Math.Max(1, size));
            try
            {
                ByteSourceStream raw = new(_source, dataOffset, entry.CompressedSize);
                using Stream content = entry.Method == 0 ? raw : new DeflateStream(raw, CompressionMode.Decompress);
                try
                {
                    content.ReadExactly(rented.AsSpan(0, size));
                }
                catch (EndOfStreamException ex)
                {
                    throw new InvalidDataException("The ZIP entry produced less data than its declared uncompressed size.", ex);
                }
                if (content.ReadByte() != -1)
                {
                    throw new InvalidDataException("The ZIP entry produced more data than its declared uncompressed size.");
                }
                return new ZipPart(rented.AsMemory(0, size), rented);
            }
            catch
            {
                ArrayPool<byte>.Shared.Return(rented);
                throw;
            }
        }

        private long ResolveDataOffset(in ZipEntryRef entry)
        {
            ThrowIfLocalHeaderOutOfRange(entry.LocalHeaderOffset);
            Span<byte> header = stackalloc byte[LocalHeaderFixedSize];
            _source.ReadExactly(entry.LocalHeaderOffset, header);
            long nameStart = ParseLocalHeader(header, entry, out int nameLength, out int extraLength);
            byte[] name = ArrayPool<byte>.Shared.Rent(Math.Max(1, nameLength));
            try
            {
                _source.ReadExactly(nameStart, name.AsSpan(0, nameLength));
                ThrowIfNameMismatch(name.AsSpan(0, nameLength), entry);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(name);
            }
            return nameStart + nameLength + extraLength;
        }

        private async ValueTask<long> ResolveDataOffsetAsync(ZipEntryRef entry, CancellationToken ct)
        {
            ThrowIfLocalHeaderOutOfRange(entry.LocalHeaderOffset);
            byte[] buffer = ArrayPool<byte>.Shared.Rent(LocalHeaderFixedSize + ushort.MaxValue);
            try
            {
                await _source.ReadExactlyAsync(entry.LocalHeaderOffset, buffer.AsMemory(0, LocalHeaderFixedSize), ct).ConfigureAwait(false);
                long nameStart = ParseLocalHeader(buffer.AsSpan(0, LocalHeaderFixedSize), entry, out int nameLength, out int extraLength);
                await _source.ReadExactlyAsync(nameStart, buffer.AsMemory(LocalHeaderFixedSize, nameLength), ct).ConfigureAwait(false);
                ThrowIfNameMismatch(buffer.AsSpan(LocalHeaderFixedSize, nameLength), entry);
                return nameStart + nameLength + extraLength;
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }

        private void ThrowIfLocalHeaderOutOfRange(long headerOffset)
        {
            if (headerOffset < 0 || headerOffset > _source.Length - LocalHeaderFixedSize)
            {
                throw new InvalidDataException("The ZIP local file header is out of range.");
            }
        }

        private long ParseLocalHeader(ReadOnlySpan<byte> header, in ZipEntryRef entry, out int nameLength, out int extraLength)
        {
            if (BinaryPrimitives.ReadInt32LittleEndian(header) != LocalFileHeaderSignature)
            {
                throw new InvalidDataException("Invalid ZIP local file header signature.");
            }
            nameLength = BinaryPrimitives.ReadUInt16LittleEndian(header[26..]);
            extraLength = BinaryPrimitives.ReadUInt16LittleEndian(header[28..]);
            long nameStart = entry.LocalHeaderOffset + LocalHeaderFixedSize;
            if (nameLength > _source.Length - nameStart)
            {
                throw new InvalidDataException("The ZIP local file header name runs past the end of the file.");
            }
            if (nameLength != entry.NameLength)
            {
                throw new InvalidDataException("The ZIP local file header name does not match the central directory.");
            }
            return nameStart;
        }

        private void ThrowIfNameMismatch(ReadOnlySpan<byte> localName, in ZipEntryRef entry)
        {
            if (!localName.SequenceEqual(_directory.Span.Slice(entry.NameStart, entry.NameLength)))
            {
                throw new InvalidDataException("The ZIP local file header name does not match the central directory.");
            }
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0)
            {
                return;
            }
            if (_entries.Length > 0)
            {
                ArrayPool<ZipEntryRef>.Shared.Return(_entries);
            }
            if (_directoryRented is not null)
            {
                ArrayPool<byte>.Shared.Return(_directoryRented);
            }
            _source.Dispose();
        }

        private static long FindEocd(ReadOnlySpan<byte> span)
        {
            if (span.Length < EocdFixedSize)
            {
                throw new InvalidDataException("End of central directory record not found.");
            }
            int windowSize = (int)Math.Min(span.Length, 65535L + EocdFixedSize);
            int windowStart = span.Length - windowSize;
            ReadOnlySpan<byte> window = span[windowStart..];
            int searchEnd = window.Length;
            while (searchEnd >= EocdFixedSize)
            {
                int index = window[..searchEnd].LastIndexOf(EocdSignatureBytes);
                if (index < 0)
                {
                    break;
                }
                if (index + EocdFixedSize > window.Length)
                {
                    searchEnd = index;
                    continue;
                }
                int commentLength = BinaryPrimitives.ReadUInt16LittleEndian(window.Slice(index + 20, 2));
                if (index + EocdFixedSize + commentLength == window.Length)
                {
                    return windowStart + index;
                }
                searchEnd = index;
            }
            throw new InvalidDataException("End of central directory record not found.");
        }

        private static (long CdOffset, long CdSize, long Count) ReadEocdRecord(ReadOnlySpan<byte> span, long eocdOffset, out long zip64EocdOffset)
        {
            ReadOnlySpan<byte> eocd = span.Slice((int)eocdOffset, EocdFixedSize);
            ushort declaredCount = BinaryPrimitives.ReadUInt16LittleEndian(eocd[10..]);
            uint cdSize = BinaryPrimitives.ReadUInt32LittleEndian(eocd[12..]);
            uint cdOffset = BinaryPrimitives.ReadUInt32LittleEndian(eocd[16..]);

            bool needsZip64 = declaredCount == 0xFFFF || cdSize == Zip64SentinelU32 || cdOffset == Zip64SentinelU32;
            if (!needsZip64 || eocdOffset < Zip64LocatorSize)
            {
                zip64EocdOffset = -1;
                return (cdOffset, cdSize, declaredCount);
            }

            ReadOnlySpan<byte> locator = span.Slice((int)(eocdOffset - Zip64LocatorSize), Zip64LocatorSize);
            if (BinaryPrimitives.ReadInt32LittleEndian(locator) != Zip64EocdLocatorSignature)
            {
                zip64EocdOffset = -1;
                return (cdOffset, cdSize, declaredCount);
            }
            zip64EocdOffset = BinaryPrimitives.ReadInt64LittleEndian(locator[8..]);
            if (zip64EocdOffset < 0)
            {
                throw new InvalidDataException("The ZIP64 end of central directory record is out of range.");
            }
            return (cdOffset, cdSize, declaredCount);
        }

        private static (long CdOffset, long CdSize, long Count) ReadZip64Eocd(ReadOnlySpan<byte> span, long offset)
        {
            if (offset < 0 || offset > span.Length - Zip64EocdFixedSize)
            {
                throw new InvalidDataException("The ZIP64 end of central directory record is out of range.");
            }
            ReadOnlySpan<byte> record = span.Slice((int)offset, Zip64EocdFixedSize);
            if (BinaryPrimitives.ReadInt32LittleEndian(record) != Zip64EocdSignature)
            {
                throw new InvalidDataException("The ZIP64 end of central directory signature is invalid.");
            }
            long count = BinaryPrimitives.ReadInt64LittleEndian(record[32..]);
            long cdSize = BinaryPrimitives.ReadInt64LittleEndian(record[40..]);
            long cdOffset = BinaryPrimitives.ReadInt64LittleEndian(record[48..]);
            return (cdOffset, cdSize, count);
        }

        private static int WalkCentralDirectory(ReadOnlySpan<byte> directory, ref ZipEntryRef[] entries, ExcelReaderOptions options)
        {
            int end = directory.Length;
            int pos = 0;
            int count = 0;
            while (pos < end)
            {
                if (end - pos < CentralDirectoryFixedSize)
                {
                    throw new InvalidDataException("The ZIP central directory is truncated.");
                }
                ReadOnlySpan<byte> record = directory.Slice(pos, CentralDirectoryFixedSize);
                if (BinaryPrimitives.ReadInt32LittleEndian(record) != CentralDirectorySignature)
                {
                    throw new InvalidDataException("Invalid ZIP central directory record signature.");
                }
                CdFixedFields fields = ParseCentralDirectoryFixed(record);
                if ((fields.Flags & EncryptedFlag) != 0)
                {
                    throw new NotSupportedException("Encrypted ZIP entries are not supported.");
                }

                int nameStart = pos + CentralDirectoryFixedSize;
                int variableLength = fields.NameLength + fields.ExtraLength + fields.CommentLength;
                if (variableLength > end - nameStart)
                {
                    throw new InvalidDataException("The ZIP central directory is truncated.");
                }

                (long compressed, long uncompressed, long localOffset) = ResolveZip64Sizes(
                    directory, nameStart + fields.NameLength, fields.ExtraLength, fields);

                EnsureCapacity(ref entries, count);
                entries[count] = new ZipEntryRef
                {
                    NameStart = nameStart,
                    NameLength = fields.NameLength,
                    LocalHeaderOffset = localOffset,
                    CompressedSize = compressed,
                    UncompressedSize = uncompressed,
                    Method = fields.Method,
                    Flags = fields.Flags,
                };
                count++;
                LimitChecks.ThrowIfTooManyEntries(count, options);
                pos = nameStart + variableLength;
            }
            return count;
        }

        [StructLayout(LayoutKind.Auto)]
        private readonly struct CdFixedFields
        {
            internal ushort Flags { get; init; }
            internal ushort Method { get; init; }
            internal ushort NameLength { get; init; }
            internal ushort ExtraLength { get; init; }
            internal ushort CommentLength { get; init; }
            internal uint CompressedSize { get; init; }
            internal uint UncompressedSize { get; init; }
            internal uint LocalHeaderOffset { get; init; }
        }

        private static CdFixedFields ParseCentralDirectoryFixed(ReadOnlySpan<byte> record)
        {
            return new CdFixedFields
            {
                Flags = BinaryPrimitives.ReadUInt16LittleEndian(record[8..]),
                Method = BinaryPrimitives.ReadUInt16LittleEndian(record[10..]),
                CompressedSize = BinaryPrimitives.ReadUInt32LittleEndian(record[20..]),
                UncompressedSize = BinaryPrimitives.ReadUInt32LittleEndian(record[24..]),
                NameLength = BinaryPrimitives.ReadUInt16LittleEndian(record[28..]),
                ExtraLength = BinaryPrimitives.ReadUInt16LittleEndian(record[30..]),
                CommentLength = BinaryPrimitives.ReadUInt16LittleEndian(record[32..]),
                LocalHeaderOffset = BinaryPrimitives.ReadUInt32LittleEndian(record[42..]),
            };
        }

        private static (long Compressed, long Uncompressed, long LocalOffset) ResolveZip64Sizes(
            ReadOnlySpan<byte> span, int extraStart, int extraLength, CdFixedFields fields)
        {
            bool needsZip64 = fields.CompressedSize == Zip64SentinelU32 || fields.UncompressedSize == Zip64SentinelU32
                || fields.LocalHeaderOffset == Zip64SentinelU32;
            if (!needsZip64)
            {
                return (fields.CompressedSize, fields.UncompressedSize, fields.LocalHeaderOffset);
            }
            if (extraLength > span.Length - extraStart)
            {
                throw new InvalidDataException("The ZIP extra field is truncated.");
            }

            ReadOnlySpan<byte> extra = span.Slice(extraStart, extraLength);
            int pos = 0;
            while (pos + 4 <= extra.Length)
            {
                ushort id = BinaryPrimitives.ReadUInt16LittleEndian(extra[pos..]);
                ushort dataSize = BinaryPrimitives.ReadUInt16LittleEndian(extra[(pos + 2)..]);
                int dataStart = pos + 4;
                if (dataStart + dataSize > extra.Length)
                {
                    throw new InvalidDataException("The ZIP64 extra field is truncated.");
                }
                if (id == Zip64ExtraFieldId)
                {
                    return ReadZip64ExtraField(extra.Slice(dataStart, dataSize), fields);
                }
                pos = dataStart + dataSize;
            }
            throw new InvalidDataException("Expected a ZIP64 extra field but none was present.");
        }

        private static (long Compressed, long Uncompressed, long LocalOffset) ReadZip64ExtraField(ReadOnlySpan<byte> data, CdFixedFields fields)
        {
            int pos = 0;
            long uncompressed = fields.UncompressedSize == Zip64SentinelU32 ? ReadNextInt64(data, ref pos) : fields.UncompressedSize;
            long compressed = fields.CompressedSize == Zip64SentinelU32 ? ReadNextInt64(data, ref pos) : fields.CompressedSize;
            long localOffset = fields.LocalHeaderOffset == Zip64SentinelU32 ? ReadNextInt64(data, ref pos) : fields.LocalHeaderOffset;
            if (uncompressed < 0 || compressed < 0 || localOffset < 0)
            {
                throw new InvalidDataException("The ZIP64 extra field holds a negative size or offset.");
            }
            return (compressed, uncompressed, localOffset);
        }

        private static long ReadNextInt64(ReadOnlySpan<byte> data, ref int pos)
        {
            if (pos + 8 > data.Length)
            {
                throw new InvalidDataException("The ZIP64 extra field is truncated.");
            }
            long value = BinaryPrimitives.ReadInt64LittleEndian(data[pos..]);
            pos += 8;
            return value;
        }

        private static void EnsureCapacity(ref ZipEntryRef[] entries, int count)
        {
            if (count < entries.Length)
            {
                return;
            }
            ZipEntryRef[] bigger = ArrayPool<ZipEntryRef>.Shared.Rent(entries.Length * 2);
            entries.AsSpan(0, count).CopyTo(bigger);
            ArrayPool<ZipEntryRef>.Shared.Return(entries);
            entries = bigger;
        }

        private static ZipPart InflateToPart(ReadOnlyMemory<byte> compressed, long uncompressedSize)
        {
            int size = checked((int)uncompressedSize);
            byte[] rented = ArrayPool<byte>.Shared.Rent(size);
            try
            {
                using MemoryStream source = ToReadableMemoryStream(compressed);
                using var inflate = new DeflateStream(source, CompressionMode.Decompress);
                try
                {
                    inflate.ReadExactly(rented.AsSpan(0, size));
                }
                catch (EndOfStreamException ex)
                {
                    throw new InvalidDataException("The ZIP entry produced less data than its declared uncompressed size.", ex);
                }
                if (inflate.ReadByte() != -1)
                {
                    throw new InvalidDataException("The ZIP entry produced more data than its declared uncompressed size.");
                }
                return new ZipPart(rented.AsMemory(0, size), rented);
            }
            catch
            {
                ArrayPool<byte>.Shared.Return(rented);
                throw;
            }
        }

        private static MemoryStream ToReadableMemoryStream(ReadOnlyMemory<byte> compressed)
        {
            if (MemoryMarshal.TryGetArray(compressed, out ArraySegment<byte> segment))
            {
                return new MemoryStream(segment.Array!, segment.Offset, segment.Count, writable: false);
            }
            return new MemoryStream(compressed.ToArray(), writable: false);
        }
    }
}
