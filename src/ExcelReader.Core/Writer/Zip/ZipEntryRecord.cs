using System.IO.Compression;
using System.Text;

namespace ExcelReader.Core.Writer.Zip
{
    internal sealed class ZipEntryRecord
    {
        internal ZipEntryRecord(string name, CompressionLevel level, long offset, DateTime modified)
        {
            Name = Encoding.UTF8.GetBytes(name);
            Level = level;
            Offset = offset;
            (Time, Date) = ZipRecords.DosDateTime(modified);
            Flags = ZipRecords.LevelFlags(level);
            if (!Ascii.IsValid(name))
            {
                Flags |= ZipRecords.FlagUtf8Name;
            }
            Method = ZipRecords.MethodDeflate;
            if (level == CompressionLevel.NoCompression)
            {
                Method = ZipRecords.MethodStored;
            }
        }

        internal byte[] Name { get; }

        internal CompressionLevel Level { get; }

        /// <summary>Absolute position of the local header in the output stream.</summary>
        internal long Offset { get; }

        internal ushort Time { get; }

        internal ushort Date { get; }

        internal ushort Flags { get; set; }

        internal ushort Method { get; set; }

        internal uint Crc { get; set; }

        internal long CompressedSize { get; set; }

        internal long UncompressedSize { get; set; }

        internal int LocalHeaderLength
        {
            get
            {
                return ZipRecords.LocalHeaderLength + Name.Length;
            }
        }

        internal bool NeedsZip64Sizes
        {
            get
            {
                return CompressedSize >= uint.MaxValue || UncompressedSize >= uint.MaxValue;
            }
        }

        internal bool NeedsZip64Offset
        {
            get
            {
                return Offset >= uint.MaxValue;
            }
        }
    }
}
