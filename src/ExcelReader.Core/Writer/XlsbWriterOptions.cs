using System.IO.Compression;

namespace ExcelReader.Core.Writer
{
    /// <summary>Configures how an XLSB package is written.</summary>
    public sealed record XlsbWriterOptions
    {
        /// <summary>Whether the workbook uses the 1904 date system instead of the default 1900 system. Defaults to <see langword="false"/>.</summary>
        public bool Date1904 { get; init; }

        /// <summary>The compression level applied to each ZIP entry. Defaults to <see cref="CompressionLevel.Fastest"/>.</summary>
        public CompressionLevel Compression { get; init; } = CompressionLevel.Fastest;

        /// <summary>Whether string cells are deduplicated into a shared string table instead of being written inline. Defaults to <see langword="false"/>.</summary>
        public bool UseSharedStrings { get; init; }

        /// <summary>
        /// Whether each sheet's deflate runs on a background thread instead of the calling thread,
        /// overlapping compression with row serialization. Defaults to <see langword="false"/>.
        /// </summary>
        /// <remarks>
        /// Mirrors <see cref="Reader.ExcelReaderOptions.PrefetchDecompression"/>'s tradeoff on the read
        /// side: worth it for single-file batch writing, not for a concurrent server workload already
        /// saturating the CPU.
        /// </remarks>
        public bool PrefetchWrite { get; init; }

        /// <summary>The default options: 1900 date system, fastest compression, inline strings, no background deflate.</summary>
        public static XlsbWriterOptions Default { get; } = new();
    }
}
