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
            return options.SniffDialect ? Applied(options, Excel.SniffCsvDialectFromFile(path)) : options;
        }

        internal static CsvReaderOptions Resolve(Stream stream, CsvReaderOptions options)
        {
            return options.SniffDialect ? Applied(options, Excel.SniffCsvDialect(stream)) : options;
        }

        internal static CsvReaderOptions Resolve(ReadOnlyMemory<byte> data, CsvReaderOptions options)
        {
            return options.SniffDialect ? Applied(options, Excel.SniffCsvDialect(data)) : options;
        }

        internal static async ValueTask<CsvReaderOptions> ResolveAsync(string path, CsvReaderOptions options, CancellationToken ct)
        {
            return options.SniffDialect
                ? Applied(options, await Excel.SniffCsvDialectFromFileAsync(path, null, ct).ConfigureAwait(false))
                : options;
        }

        internal static async ValueTask<CsvReaderOptions> ResolveAsync(Stream stream, CsvReaderOptions options, CancellationToken ct)
        {
            return options.SniffDialect
                ? Applied(options, await Excel.SniffCsvDialectAsync(stream, null, ct).ConfigureAwait(false))
                : options;
        }

        private static CsvReaderOptions Applied(CsvReaderOptions options, CsvDialect dialect)
        {
            CsvReaderOptions applied = options.WithDialect(dialect) with { SniffDialect = false };
            return dialect.Encoding is null ? applied with { Encoding = options.Encoding } : applied;
        }
    }
}
