using System.Buffers.Binary;
using System.IO.Compression;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Zip
{
    /// <summary>
    /// The ZIP records of PKWARE's APPNOTE.TXT section 4.3, with the field values <c>ZipArchive</c>
    /// writes in create mode, and Zip64 only where a 32-bit field overflows.
    /// </summary>
    internal static class ZipRecords
    {
        internal const int LocalHeaderLength = 30;
        internal const int LocalVersionOffset = 4;
        internal const int LocalCrcOffset = 14;
        internal const ushort FlagDataDescriptor = 0x0008;
        internal const ushort FlagUtf8Name = 0x0800;
        internal const ushort MethodStored = 0;
        internal const ushort MethodDeflate = 8;
        internal const ushort VersionDefault = 20;
        internal const ushort VersionZip64 = 45;

        private const uint LocalHeaderSignature = 0x0403_4B50;
        private const uint DataDescriptorSignature = 0x0807_4B50;
        private const uint CentralHeaderSignature = 0x0201_4B50;
        private const uint Zip64EndSignature = 0x0606_4B50;
        private const uint Zip64LocatorSignature = 0x0706_4B50;
        private const uint EndSignature = 0x0605_4B50;
        private const ushort VersionMadeBy = 0x0014;
        private const ushort Zip64ExtraId = 0x0001;
        private const ulong Zip64EndRecordSize = 44;
        private const int MaxClassicCount = 0xFFFF;

        internal static ushort LevelFlags(CompressionLevel level)
        {
            switch (level)
            {
                case CompressionLevel.Fastest:
                    return 0x0006;
                case CompressionLevel.SmallestSize:
                    return 0x0002;
                default:
                    return 0;
            }
        }

        internal static (ushort Time, ushort Date) DosDateTime(DateTime value)
        {
            DateTime clamped = value;
            if (clamped.Year < 1980)
            {
                clamped = new DateTime(1980, 1, 1, 0, 0, 0, DateTimeKind.Local);
            }
            if (clamped.Year > 2107)
            {
                clamped = new DateTime(2107, 12, 31, 23, 59, 58, DateTimeKind.Local);
            }
            ushort time = (ushort)((clamped.Hour << 11) | (clamped.Minute << 5) | (clamped.Second / 2));
            ushort date = (ushort)(((clamped.Year - 1980) << 9) | (clamped.Month << 5) | clamped.Day);
            return (time, date);
        }

        internal static void WriteLocalHeader(BiffBuffer buffer, ZipEntryRecord entry)
        {
            buffer.WriteU32(LocalHeaderSignature);
            buffer.WriteU16(VersionDefault);
            buffer.WriteU16(entry.Flags);
            buffer.WriteU16(entry.Method);
            buffer.WriteU16(entry.Time);
            buffer.WriteU16(entry.Date);
            buffer.WriteU32(0);
            buffer.WriteU32(0);
            buffer.WriteU32(0);
            buffer.WriteU16(entry.Name.Length);
            buffer.WriteU16(0);
            buffer.Write(entry.Name);
        }

        /// <summary>The 12 bytes at <see cref="LocalCrcOffset"/>: CRC-32 and 32-bit sizes.</summary>
        internal static byte[] CrcAndSizes(ZipEntryRecord entry)
        {
            byte[] bytes = new byte[12];
            BinaryPrimitives.WriteUInt32LittleEndian(bytes, entry.Crc);
            BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(4), (uint)entry.CompressedSize);
            BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(8), (uint)entry.UncompressedSize);
            return bytes;
        }

        /// <summary>The 4 bytes at <see cref="LocalVersionOffset"/>: version needed and flags.</summary>
        internal static byte[] VersionAndFlags(ZipEntryRecord entry)
        {
            byte[] bytes = new byte[4];
            BinaryPrimitives.WriteUInt16LittleEndian(bytes, VersionZip64);
            BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(2), entry.Flags);
            return bytes;
        }

        internal static void WriteDataDescriptor(BiffBuffer buffer, ZipEntryRecord entry)
        {
            buffer.WriteU32(DataDescriptorSignature);
            buffer.WriteU32(entry.Crc);
            if (entry.NeedsZip64Sizes)
            {
                WriteU64(buffer, (ulong)entry.CompressedSize);
                WriteU64(buffer, (ulong)entry.UncompressedSize);
                return;
            }
            buffer.WriteU32((uint)entry.CompressedSize);
            buffer.WriteU32((uint)entry.UncompressedSize);
        }

        internal static void WriteCentralHeader(BiffBuffer buffer, ZipEntryRecord entry)
        {
            bool zip64Sizes = entry.NeedsZip64Sizes;
            bool zip64Offset = entry.NeedsZip64Offset;
            int extraFields = (zip64Sizes ? 2 : 0) + (zip64Offset ? 1 : 0);
            int extraLength = extraFields == 0 ? 0 : 4 + (8 * extraFields);
            buffer.WriteU32(CentralHeaderSignature);
            buffer.WriteU16(VersionMadeBy);
            buffer.WriteU16(extraFields == 0 ? VersionDefault : VersionZip64);
            buffer.WriteU16(entry.Flags);
            buffer.WriteU16(entry.Method);
            buffer.WriteU16(entry.Time);
            buffer.WriteU16(entry.Date);
            buffer.WriteU32(entry.Crc);
            buffer.WriteU32(Field(entry.CompressedSize, zip64Sizes));
            buffer.WriteU32(Field(entry.UncompressedSize, zip64Sizes));
            buffer.WriteU16(entry.Name.Length);
            buffer.WriteU16(extraLength);
            buffer.WriteU16(0);
            buffer.WriteU16(0);
            buffer.WriteU16(0);
            buffer.WriteU32(0);
            buffer.WriteU32(Field(entry.Offset, zip64Offset));
            buffer.Write(entry.Name);
            if (extraFields == 0)
            {
                return;
            }
            buffer.WriteU16(Zip64ExtraId);
            buffer.WriteU16(8 * extraFields);
            if (zip64Sizes)
            {
                WriteU64(buffer, (ulong)entry.UncompressedSize);
                WriteU64(buffer, (ulong)entry.CompressedSize);
            }
            if (zip64Offset)
            {
                WriteU64(buffer, (ulong)entry.Offset);
            }
        }

        internal static void WriteEnd(BiffBuffer buffer, int count, long directoryOffset, long directoryLength)
        {
            bool zip64 = count >= MaxClassicCount || directoryOffset >= uint.MaxValue || directoryLength >= uint.MaxValue;
            if (zip64)
            {
                WriteZip64End(buffer, count, directoryOffset, directoryLength);
            }
            buffer.WriteU32(EndSignature);
            buffer.WriteU16(0);
            buffer.WriteU16(0);
            buffer.WriteU16(Math.Min(count, MaxClassicCount));
            buffer.WriteU16(Math.Min(count, MaxClassicCount));
            buffer.WriteU32(Field(directoryLength, directoryLength >= uint.MaxValue));
            buffer.WriteU32(Field(directoryOffset, directoryOffset >= uint.MaxValue));
            buffer.WriteU16(0);
        }

        private static void WriteZip64End(BiffBuffer buffer, int count, long directoryOffset, long directoryLength)
        {
            long zip64EndOffset = directoryOffset + directoryLength;
            buffer.WriteU32(Zip64EndSignature);
            WriteU64(buffer, Zip64EndRecordSize);
            buffer.WriteU16(VersionZip64);
            buffer.WriteU16(VersionZip64);
            buffer.WriteU32(0);
            buffer.WriteU32(0);
            WriteU64(buffer, (ulong)count);
            WriteU64(buffer, (ulong)count);
            WriteU64(buffer, (ulong)directoryLength);
            WriteU64(buffer, (ulong)directoryOffset);
            buffer.WriteU32(Zip64LocatorSignature);
            buffer.WriteU32(0);
            WriteU64(buffer, (ulong)zip64EndOffset);
            buffer.WriteU32(1);
        }

        private static uint Field(long value, bool overflows)
        {
            if (overflows)
            {
                return uint.MaxValue;
            }
            return (uint)value;
        }

        private static void WriteU64(BiffBuffer buffer, ulong value)
        {
            BinaryPrimitives.WriteUInt64LittleEndian(buffer.GetSpan(sizeof(ulong)), value);
            buffer.Advance(sizeof(ulong));
        }
    }
}
