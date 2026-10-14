using System.IO.Compression;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Zip
{
    /// <summary>
    /// Compresses one run of bytes into a deflate segment: data that ends on a byte boundary and has
    /// no final block, so segments laid end to end, followed by <see cref="FinalBlock"/>, are one
    /// valid deflate stream. A segment never refers to bytes outside itself.
    /// </summary>
    internal static class SegmentCompressor
    {
        /// <summary>An empty fixed-Huffman block marked final: it ends a stream of segments.</summary>
        internal static ReadOnlySpan<byte> FinalBlock
        {
            get
            {
                return [0x03, 0x00];
            }
        }

        internal static void Compress(ReadOnlySpan<byte> input, CompressionLevel level, BiffBuffer output)
        {
            if (level == CompressionLevel.NoCompression)
            {
                output.Write(input);
                return;
            }
            // DeflateEncoder writes and compares native-order words, which is only deflate's bit order on little-endian.
            if (level == CompressionLevel.Fastest && BitConverter.IsLittleEndian)
            {
                using var encoder = new DeflateEncoder();
                encoder.Write(input, output);
                encoder.FlushSync(output);
                return;
            }
            CompressWithZlib(input, level, output);
        }

        // DeflateStream.Flush is zlib's sync flush, which ends on a byte boundary; the final block that
        // Dispose writes afterwards is dropped.
        private static void CompressWithZlib(ReadOnlySpan<byte> input, CompressionLevel level, BiffBuffer output)
        {
            var sink = new BufferSink(output);
            using var zlib = new DeflateStream(sink, level);
            zlib.Write(input);
            zlib.Flush();
            sink.Discard = true;
        }
    }
}
