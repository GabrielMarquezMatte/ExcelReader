using System.Buffers;
using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Runtime.InteropServices;

namespace ExcelReader.Core.Reader
{
    // Offsets into the whole-file buffer, not allocated strings, so building the index never
    // allocates one object per ZIP entry.
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

    // A stored entry aliases the caller's buffer (Rented is null, nothing to return). A deflated
    // entry owns a pooled array that must be returned via Dispose.
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

    // Reads a ZIP central directory directly out of a fully-materialized, in-memory archive: no
    // ZipArchive, no per-entry allocation for the parts the caller never looks up. Every offset and
    // length is bounds-checked against the file before use, since the input is untrusted.
    internal sealed class ZipMemoryIndex : IDisposable
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

        private readonly ReadOnlyMemory<byte> _file;
        private readonly ZipEntryRef[] _entries;
        private bool _disposed;

        private ZipMemoryIndex(ReadOnlyMemory<byte> file, ZipEntryRef[] entries, int count)
        {
            _file = file;
            _entries = entries;
            Count = count;
        }

        internal int Count { get; }

        internal static ZipMemoryIndex Create(ReadOnlyMemory<byte> file, ExcelReaderOptions options)
        {
            ReadOnlySpan<byte> span = file.Span;
            long eocdOffset = FindEocd(span);
            (long cdOffset, long cdSize, long declaredCount) = ReadEocdRecord(span, eocdOffset, out long zip64EocdOffset);
            if (zip64EocdOffset >= 0)
            {
                (cdOffset, cdSize, declaredCount) = ReadZip64Eocd(span, zip64EocdOffset);
            }
            // cdOffset and cdSize are both attacker-controlled longs (straight from the ZIP64 EOCD when
            // present); `cdOffset + cdSize` can overflow and wrap negative for a value near long.MaxValue,
            // making this check pass when it should reject. Comparing via subtraction instead can never
            // overflow: cdOffset is already bounded >= 0 here, and span.Length (an int) minus cdSize can
            // go very negative for a huge cdSize but never wraps, so the comparison still rejects it.
            if (cdOffset < 0 || cdSize < 0 || cdOffset > span.Length - cdSize)
            {
                throw new InvalidDataException("The ZIP central directory is out of range.");
            }

            // Math.Max keeps the lower bound (16) from ever exceeding the upper bound: a caller can
            // legitimately configure MaxZipEntries below 16, and Math.Clamp throws if min > max.
            long maxHint = Math.Max(16, options.MaxZipEntries > 0 ? options.MaxZipEntries : 65_536);
            ZipEntryRef[] entries = ArrayPool<ZipEntryRef>.Shared.Rent((int)Math.Clamp(declaredCount, 16, maxHint));
            int count;
            try
            {
                count = WalkCentralDirectory(span, cdOffset, cdSize, ref entries, options);
            }
            catch
            {
                ArrayPool<ZipEntryRef>.Shared.Return(entries);
                throw;
            }
            return new ZipMemoryIndex(file, entries, count);
        }

        internal bool TryGetEntry(ReadOnlySpan<byte> utf8Name, out ZipEntryRef entry)
        {
            ReadOnlySpan<byte> fileSpan = _file.Span;
            foreach (ref readonly ZipEntryRef candidate in _entries.AsSpan(0, Count))
            {
                if (fileSpan.Slice(candidate.NameStart, candidate.NameLength).SequenceEqual(utf8Name))
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
            LimitChecks.ThrowIfEntryLengthExceeds(entry.UncompressedSize, counter.Remaining, nameof(ExcelReaderOptions.MaxTotalDecompressedBytes));
            if (entryLimit > 0)
            {
                LimitChecks.ThrowIfEntryLengthExceeds(entry.UncompressedSize, entryLimit, entryLimitName);
            }
            if (entry.UncompressedSize > Array.MaxLength)
            {
                throw new ExcelLimitExceededException("ArrayMaxLength", Array.MaxLength, entry.UncompressedSize);
            }

            ReadOnlyMemory<byte> compressed = ResolveCompressedSlice(entry);
            ZipPart part = entry.Method switch
            {
                0 => new ZipPart(compressed, rented: null),
                8 => InflateToPart(compressed, entry.UncompressedSize),
                _ => throw new NotSupportedException($"Unsupported ZIP compression method: {entry.Method}."),
            };
            counter.Add(entry.UncompressedSize);
            return part;
        }

        // default(ZipPart) (empty Memory, nothing to return on Dispose) stands in for a missing part —
        // the same convention ZipEntryBytes.Read uses for a missing entry on the streamed path.
        internal ZipPart OpenPartOrDefault(ReadOnlySpan<byte> utf8Name, DecompressedByteCounter counter)
        {
            return TryGetEntry(utf8Name, out ZipEntryRef entry) ? OpenPart(entry, counter) : default;
        }

        // In-memory ZIP path's worksheet entry point: opens a Stream over the
        // entry's compressed bytes instead of eagerly materializing the whole part, so the caller (the
        // XlsxReader/XlsbReader enumerator) can reuse the exact same PrefetchStream/LimitedReadStream
        // pipeline as the ZipArchive path (WorkbookLookups.Wrap) and overlap inflate with row parsing.
        // Declared-size limits are enforced incrementally as bytes are actually read, matching the
        // streamed path, rather than eagerly against entry.UncompressedSize the way OpenPart must.
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "Ownership of the opened entry stream and its wrapper transfers to the caller.")]
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Ownership of the opened entry stream and its wrapper transfers to the caller.")]
        internal Stream OpenEntryStream(in ZipEntryRef entry, DecompressedByteCounter counter, ExcelReaderOptions options, string entryLimitName = "", long entryLimit = 0)
        {
            ReadOnlyMemory<byte> compressed = ResolveCompressedSlice(entry);
            Stream opened = entry.Method switch
            {
                0 => ToReadableMemoryStream(compressed),
                8 => new DeflateStream(ToReadableMemoryStream(compressed), CompressionMode.Decompress),
                _ => throw new NotSupportedException($"Unsupported ZIP compression method: {entry.Method}."),
            };
            return WorkbookLookups.Wrap(opened, counter, options, entryLimitName, entryLimit, entry.UncompressedSize);
        }

        // Bounds-checks and slices the entry's compressed bytes out of the whole-file buffer. Shared by
        // OpenPart (which decompresses eagerly into a ZipPart) and OpenEntryStream (which hands the
        // slice to a DeflateStream instead). CompressedSize must fit an int for the Slice cast below
        // regardless of which caller uses it, so that check lives here rather than in each caller.
        private ReadOnlyMemory<byte> ResolveCompressedSlice(in ZipEntryRef entry)
        {
            if (entry.CompressedSize > Array.MaxLength)
            {
                throw new ExcelLimitExceededException("ArrayMaxLength", Array.MaxLength, entry.CompressedSize);
            }
            long dataOffset = ResolveDataOffset(entry);
            ReadOnlySpan<byte> fileSpan = _file.Span;
            if (dataOffset < 0 || dataOffset + entry.CompressedSize > fileSpan.Length)
            {
                throw new InvalidDataException("The ZIP entry data runs past the end of the file.");
            }
            return _file.Slice((int)dataOffset, (int)entry.CompressedSize);
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            if (_entries.Length > 0)
            {
                ArrayPool<ZipEntryRef>.Shared.Return(_entries);
            }
        }

        // Scans backward for the last signature occurrence whose declared comment length reaches
        // exactly to the end of the file — a ZIP-like byte sequence inside an earlier part's content
        // (e.g. an embedded image) must not be mistaken for the real record.
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

        // zip64EocdOffset stays -1 unless a sentinel field forces a real ZIP64 lookup, so Create only
        // pays for the extra 20+56 byte reads when the archive actually needs them.
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
            return (cdOffset, cdSize, declaredCount);
        }

        private static (long CdOffset, long CdSize, long Count) ReadZip64Eocd(ReadOnlySpan<byte> span, long offset)
        {
            // SEC-6: offset comes straight from the ZIP64 locator's 8-byte field, so a value near
            // long.MaxValue makes `offset + Zip64EocdFixedSize` overflow and wrap negative, passing this
            // check when it should reject — then `(int)offset` truncates to an arbitrary value and
            // Slice either throws a raw ArgumentOutOfRangeException or, worse, succeeds on the wrong
            // range. The subtraction form below can't overflow (span.Length is a small int) and rejects
            // the same inputs correctly.
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

        private static int WalkCentralDirectory(ReadOnlySpan<byte> span, long cdOffset, long cdSize, ref ZipEntryRef[] entries, ExcelReaderOptions options)
        {
            long end = cdOffset + cdSize;
            long pos = cdOffset;
            int count = 0;
            while (pos < end)
            {
                if (pos + CentralDirectoryFixedSize > span.Length)
                {
                    throw new InvalidDataException("The ZIP central directory is truncated.");
                }
                ReadOnlySpan<byte> record = span.Slice((int)pos, CentralDirectoryFixedSize);
                if (BinaryPrimitives.ReadInt32LittleEndian(record) != CentralDirectorySignature)
                {
                    throw new InvalidDataException("Invalid ZIP central directory record signature.");
                }
                CdFixedFields fields = ParseCentralDirectoryFixed(record);
                if ((fields.Flags & EncryptedFlag) != 0)
                {
                    throw new NotSupportedException("Encrypted ZIP entries are not supported.");
                }

                long nameStart = pos + CentralDirectoryFixedSize;
                long recordEnd = nameStart + fields.NameLength + fields.ExtraLength + fields.CommentLength;
                if (recordEnd > span.Length || recordEnd > end)
                {
                    throw new InvalidDataException("The ZIP central directory is truncated.");
                }

                (long compressed, long uncompressed, long localOffset) = ResolveZip64Sizes(
                    span, (int)(nameStart + fields.NameLength), fields.ExtraLength, fields);

                EnsureCapacity(ref entries, count);
                entries[count] = new ZipEntryRef
                {
                    NameStart = (int)nameStart,
                    NameLength = fields.NameLength,
                    LocalHeaderOffset = localOffset,
                    CompressedSize = compressed,
                    UncompressedSize = uncompressed,
                    Method = fields.Method,
                    Flags = fields.Flags,
                };
                count++;
                // SEC-6: was previously enforced only after the whole central directory had been
                // walked and every entry materialized into `entries` — a huge malicious CD count still
                // paid the full walk/grow cost before rejection. Check as each entry lands instead.
                LimitChecks.ThrowIfTooManyEntries(count, options);
                pos = recordEnd;
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

        // Only entries whose 32-bit field is the ZIP64 sentinel actually carry a replacement value in
        // the extra field, and they appear in this fixed order (uncompressed, compressed, local
        // offset, disk) — reading anything not sentineled would misalign every field after it.
        private static (long Compressed, long Uncompressed, long LocalOffset) ResolveZip64Sizes(
            ReadOnlySpan<byte> span, int extraStart, int extraLength, CdFixedFields fields)
        {
            bool needsZip64 = fields.CompressedSize == Zip64SentinelU32 || fields.UncompressedSize == Zip64SentinelU32
                || fields.LocalHeaderOffset == Zip64SentinelU32;
            if (!needsZip64)
            {
                return (fields.CompressedSize, fields.UncompressedSize, fields.LocalHeaderOffset);
            }
            if (extraStart + extraLength > span.Length)
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

        private long ResolveDataOffset(in ZipEntryRef entry)
        {
            ReadOnlySpan<byte> fileSpan = _file.Span;
            long headerOffset = entry.LocalHeaderOffset;
            if (headerOffset < 0 || headerOffset + LocalHeaderFixedSize > fileSpan.Length)
            {
                throw new InvalidDataException("The ZIP local file header is out of range.");
            }
            ReadOnlySpan<byte> header = fileSpan.Slice((int)headerOffset, LocalHeaderFixedSize);
            if (BinaryPrimitives.ReadInt32LittleEndian(header) != LocalFileHeaderSignature)
            {
                throw new InvalidDataException("Invalid ZIP local file header signature.");
            }
            ushort nameLength = BinaryPrimitives.ReadUInt16LittleEndian(header[26..]);
            ushort extraLength = BinaryPrimitives.ReadUInt16LittleEndian(header[28..]);
            return headerOffset + LocalHeaderFixedSize + nameLength + extraLength;
        }

        // This is a known stopgap: it round-trips through stdlib DeflateStream, which needs an
        // array-backed ReadOnlyMemory<byte> (or a one-time copy). A zero-allocation span-based
        // inflater would remove that, but is a separate, larger, not-yet-started follow-up.
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

        // MemoryStream needs an array; when `compressed` already aliases one (the common case — the
        // caller's whole-file buffer), wrap it with zero copies. A non-array-backed
        // ReadOnlyMemory<byte> (e.g. a custom MemoryManager<byte>) is the rare fallback that pays one
        // copy here; Phase 2 removes the need for this entirely.
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
