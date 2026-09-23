using System.IO.Compression;

namespace ExcelReader.Core.Writer
{
    /// <summary>Configures how an XLSX package is written.</summary>
    public sealed record XlsxWriterOptions
    {
        /// <summary>The compression level applied to each ZIP entry. Defaults to <see cref="CompressionLevel.Fastest"/>.</summary>
        public CompressionLevel Compression { get; init; } = CompressionLevel.Fastest;

        /// <summary>Whether string cells are deduplicated into a shared string table instead of being inlined. Defaults to <see langword="false"/>.</summary>
        public bool UseSharedStrings { get; init; }

        /// <summary>
        /// Whether each sheet's deflate runs on a background thread instead of the calling thread,
        /// overlapping compression with row serialization. Defaults to <see langword="false"/>.
        /// </summary>
        /// <remarks>
        /// Intended for single-file batch writing, where overlapping deflate with row-building shortens
        /// one write's wall-clock time — not for concurrent server workloads, where the extra background
        /// thread per sheet competes with work already saturating the CPU (mirrors
        /// <see cref="Reader.ExcelReaderOptions.PrefetchDecompression"/>'s tradeoff on the read side).
        /// </remarks>
        public bool PrefetchWrite { get; init; }

        /// <summary>The default options: fastest compression, inline strings, no background deflate.</summary>
        public static XlsxWriterOptions Default { get; } = new();
    }
}
