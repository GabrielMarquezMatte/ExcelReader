using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal static class ParallelCsvFactory
    {
        private const long MinParallelBytes = 4 * 64 * 1024;

        [SuppressMessage("Performance", "CA1849:Call async methods when in an async method",
            Justification = "This factory is synchronous by design (see the VSTHRD200 suppression above); opening the file here, once, before any enumeration starts, is deliberate rather than a blocking call inside an async method.")]
        internal static IAsyncEnumerable<T> Create<T>(
            string path, int degreeOfParallelism, CsvReaderOptions? readerOptions, ExcelParser<T> parser, CancellationToken ct)
        {
            CsvReaderOptions options = CsvDialectResolver.Resolve(path, readerOptions ?? CsvReaderOptions.Default);
            int dop = Normalize(degreeOfParallelism);

            var info = new FileInfo(path);
            if (!CanPartition(dop, info.Length, options))
            {
                return Sequential(Excel.FromCsvFile(path, options), parser, ct);
            }

            SafeFileHandle handle = File.OpenHandle(path, FileMode.Open, FileAccess.Read, FileShare.Read, FileOptions.Asynchronous | FileOptions.RandomAccess);
            return Build(new CsvChunkSource(handle, info.Length), options, parser, dop, chunkSizeOverride: 0, ownedHandle: handle, ct);
        }

        internal static IAsyncEnumerable<T> Create<T>(
            ReadOnlyMemory<byte> data, int degreeOfParallelism, CsvReaderOptions? readerOptions, ExcelParser<T> parser, CancellationToken ct)
        {
            CsvReaderOptions options = CsvDialectResolver.Resolve(data, readerOptions ?? CsvReaderOptions.Default);
            int dop = Normalize(degreeOfParallelism);

            if (!CanPartition(dop, data.Length, options))
            {
                return Sequential(Excel.FromCsv(data, options), parser, ct);
            }
            return Build(new CsvChunkSource(data), options, parser, dop, chunkSizeOverride: 0, ownedHandle: null, ct);
        }

        [SuppressMessage("Performance", "CA1849:Call async methods when in an async method",
            Justification = "This factory is synchronous by design (see the VSTHRD200 suppression above); wrapping the caller-owned stream here, once, before any enumeration starts, is deliberate rather than a blocking call inside an async method.")]
        internal static IAsyncEnumerable<T> Create<T>(
            Stream stream, int degreeOfParallelism, CsvReaderOptions? readerOptions, ExcelParser<T> parser, CancellationToken ct)
        {
            CsvReaderOptions options = CsvDialectResolver.Resolve(stream, readerOptions ?? CsvReaderOptions.Default);
            int dop = Normalize(degreeOfParallelism);

            if (!CsvSourceResolver.TryResolve(stream, out CsvChunkSource source)
                || !CanPartition(dop, source.Length, options))
            {
                return Sequential(Excel.FromCsv(stream, leaveOpen: true, options), parser, ct);
            }
            return Build(source, options, parser, dop, chunkSizeOverride: 0, ownedHandle: null, ct);
        }

        internal static IAsyncEnumerable<T> CreateWithChunkSize<T>(
            ReadOnlyMemory<byte> data, int degreeOfParallelism, int chunkSize, CsvReaderOptions? readerOptions, ExcelParser<T> parser, CancellationToken ct)
        {
            CsvReaderOptions options = CsvDialectResolver.Resolve(data, readerOptions ?? CsvReaderOptions.Default);
            return Build(new CsvChunkSource(data), options, parser, Normalize(degreeOfParallelism), chunkSize, ownedHandle: null, ct);
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
            SafeFileHandle? ownedHandle,
            CancellationToken ct)
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
                ownedHandle?.Dispose();
                return Empty<T>();
            }

            CsvChunkPlan plan = CsvChunkPlan.Create(dataStart, source.Length - dataStart, dop, chunkSizeOverride);
            if (plan.Count <= 1)
            {
                ownedHandle?.Dispose();
                return Sequential(source.OpenReader(options), parser, ct);
            }

            return new ParallelCsvEnumerable<T>(source, plan, dataStart, map, info, options, config, dop, ownedHandle);
        }

        private static CsvEnumerable<T> Sequential<T>(
            CsvReader reader, ExcelParser<T> parser, CancellationToken ct)
        {
            return new CsvEnumerable<T>(reader, parser.Config, parser.CsvInfo, ownsReader: true, ct);
        }

        private static async IAsyncEnumerable<T> Empty<T>()
        {
            await Task.CompletedTask.ConfigureAwait(false);
            yield break;
        }
    }
}
