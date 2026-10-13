using System.IO.Compression;
using ExcelReader.Core.Writer.Internal;
using ExcelReader.Core.Writer.Zip;

namespace ExcelReader.Core.Writer.Parallel
{
    /// <summary>
    /// One batch of rows after formatting and compression: the deflate segment, the CRC and length of
    /// the rows it was made from, and the range of rows and columns they used.
    /// </summary>
    internal sealed class BatchSegment : IDisposable
    {
        internal BatchSegment(int rows)
        {
            Rows = rows;
        }

        internal BiffBuffer Data { get; } = new(1 << 16);

        internal int Rows { get; }

        internal uint Crc { get; private set; }

        internal long RawLength { get; private set; }

        internal int RowFirst { get; set; } = -1;

        internal int RowLast { get; set; } = -1;

        internal int ColumnFirst { get; set; } = int.MaxValue;

        internal int ColumnLast { get; set; } = -1;

        internal void Fill(ReadOnlySpan<byte> raw, CompressionLevel level)
        {
            Crc = Crc32.Append(0, raw);
            RawLength = raw.Length;
            SegmentCompressor.Compress(raw, level, Data);
        }

        public void Dispose()
        {
            Data.Dispose();
        }
    }
}
