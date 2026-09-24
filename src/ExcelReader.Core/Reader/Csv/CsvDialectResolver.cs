namespace ExcelReader.Core.Reader.Csv
{
    /// <summary>
    /// Turns <see cref="CsvReaderOptions.SniffDialect"/> into a concrete dialect at the entry point that
    /// owns the source, so everything downstream — enumerators, chunk workers — sees plain options.
    /// </summary>
    internal static class CsvDialectResolver
    {
        internal static CsvReaderOptions Resolve(string path, CsvReaderOptions options)
        {
            return options.SniffDialect ? Applied(options, CsvSniffer.DetectFile(path)) : options;
        }

        internal static CsvReaderOptions Resolve(Stream stream, CsvReaderOptions options)
        {
            return options.SniffDialect ? Applied(options, CsvSniffer.Detect(stream)) : options;
        }

        internal static CsvReaderOptions Resolve(ReadOnlyMemory<byte> data, CsvReaderOptions options)
        {
            return options.SniffDialect ? Applied(options, CsvSniffer.Detect(data.Span)) : options;
        }

        internal static async ValueTask<CsvReaderOptions> ResolveAsync(string path, CsvReaderOptions options, CancellationToken ct)
        {
            return options.SniffDialect
                ? Applied(options, await CsvSniffer.DetectFileAsync(path, null, ct).ConfigureAwait(false))
                : options;
        }

        internal static async ValueTask<CsvReaderOptions> ResolveAsync(Stream stream, CsvReaderOptions options, CancellationToken ct)
        {
            return options.SniffDialect
                ? Applied(options, await CsvSniffer.DetectAsync(stream, null, ct).ConfigureAwait(false))
                : options;
        }

        private static CsvReaderOptions Applied(CsvReaderOptions options, CsvDialect dialect)
        {
            CsvReaderOptions applied = options.WithDialect(dialect) with { SniffDialect = false };
            return dialect.Encoding is null ? applied with { Encoding = options.Encoding } : applied;
        }
    }
}
