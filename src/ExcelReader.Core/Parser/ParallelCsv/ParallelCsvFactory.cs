using System.Text;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal static class ParallelCsvFactory
    {
        private const long MinParallelBytes = 4 * 64 * 1024;

        internal static IAsyncEnumerable<T> CreateWithChunkSize<T>(
            ReadOnlyMemory<byte> data, int degreeOfParallelism, int chunkSize, CsvReaderOptions? readerOptions, ExcelParser<T> parser, CancellationToken ct)
        {
            CsvReaderOptions options = CsvDialectResolver.Resolve(data, readerOptions ?? CsvReaderOptions.Default);
            return Build(new CsvChunkSource(data), options, parser, Normalize(degreeOfParallelism), chunkSize, ct);
        }

        internal static IAsyncEnumerable<T> Create<T>(CsvReader reader, int degreeOfParallelism, ExcelParser<T> parser, CancellationToken ct)
        {
            int dop = Normalize(degreeOfParallelism);
            if (!reader.TryGetChunkSource(out CsvChunkSource source) || !CanPartition(dop, source.Length, reader.Options))
            {
                return new CsvEnumerable<T>(reader.FirstSheet, parser.Config, parser.CsvInfo, ownedReader: null, ct);
            }
            return Build(source, reader.Options, parser, dop, chunkSizeOverride: 0, ct, reader.Lifetime, reader);
        }

        internal static int Normalize(int degreeOfParallelism)
        {
            return degreeOfParallelism == 0 ? Environment.ProcessorCount : degreeOfParallelism;
        }

        internal static bool CanPartition(int dop, long length, CsvReaderOptions options)
        {
            if (dop <= 1 || length < MinParallelBytes)
            {
                return false;
            }
            return options.Encoding is null || options.Encoding.CodePage == Encoding.UTF8.CodePage;
        }

        private static IAsyncEnumerable<T> Build<T>(
            CsvChunkSource source,
            CsvReaderOptions options,
            ExcelParser<T> parser,
            int dop,
            int chunkSizeOverride,
            CancellationToken ct,
            ReaderLifetime? lease = null,
            object? owner = null)
        {
            ExcelParserConfig config = parser.Config;
            TypeMapInfo<T> info = parser.CsvInfo;
            CsvBoundColumnMap<T> map;
            long dataStart;
            using (CsvReader headerReader = source.OpenReader(options))
            {
                map = CsvHeaderBinder.Bind(headerReader, config, info, out dataStart);
            }

            if (dataStart >= source.Length)
            {
                return Empty<T>();
            }

            CsvChunkPlan plan = CsvChunkPlan.Create(dataStart, source.Length - dataStart, dop, chunkSizeOverride);
            if (plan.Count <= 1)
            {
                return Sequential(source.OpenReader(options), parser, ct);
            }

            return new ParallelCsvEnumerable<T>(source, plan, dataStart, map, info, options, config, dop, null, lease, owner);
        }

        private static CsvEnumerable<T> Sequential<T>(
            CsvReader reader, ExcelParser<T> parser, CancellationToken ct)
        {
            return new CsvEnumerable<T>(reader.Sheets[0], parser.Config, parser.CsvInfo, ownedReader: reader, ct);
        }

        private static async IAsyncEnumerable<T> Empty<T>()
        {
            await Task.CompletedTask.ConfigureAwait(false);
            yield break;
        }
    }
}
