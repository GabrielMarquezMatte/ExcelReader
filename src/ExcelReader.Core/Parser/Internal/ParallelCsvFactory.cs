using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Reader;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Parser.Internal
{
    internal static class ParallelCsvFactory
    {
        private const long MinParallelBytes = 4 * 64 * 1024;

        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        [SuppressMessage("Performance", "CA1849:Call async methods when in an async method",
            Justification = "This factory is synchronous by design (see the VSTHRD200 suppression above); opening the file here, once, before any enumeration starts, is deliberate rather than a blocking call inside an async method.")]
        internal static IAsyncEnumerable<T> Create<T>(
            string path, int degreeOfParallelism, CsvReaderOptions? readerOptions, ExcelParserConfig? config, CancellationToken ct)
        {
            CsvReaderOptions options = readerOptions ?? CsvReaderOptions.Default;
            ExcelParserConfig parserConfig = config ?? new ExcelParserConfig();
            int dop = Normalize(degreeOfParallelism);

            var info = new FileInfo(path);
            if (!CanPartition(dop, info.Length, options))
            {
                return Sequential<T>(Excel.FromCsvFile(path, options), parserConfig, ct);
            }

            SafeFileHandle handle = File.OpenHandle(path, FileMode.Open, FileAccess.Read, FileShare.Read, FileOptions.Asynchronous | FileOptions.RandomAccess);
            return Build<T>(new CsvChunkSource(handle, info.Length), options, parserConfig, dop, chunkSizeOverride: 0, ownedHandle: handle, ct);
        }

        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        internal static IAsyncEnumerable<T> Create<T>(
            ReadOnlyMemory<byte> data, int degreeOfParallelism, CsvReaderOptions? readerOptions, ExcelParserConfig? config, CancellationToken ct)
        {
            CsvReaderOptions options = readerOptions ?? CsvReaderOptions.Default;
            ExcelParserConfig parserConfig = config ?? new ExcelParserConfig();
            int dop = Normalize(degreeOfParallelism);

            if (!CanPartition(dop, data.Length, options))
            {
                return Sequential<T>(Excel.FromCsv(data, options), parserConfig, ct);
            }
            return Build<T>(new CsvChunkSource(data), options, parserConfig, dop, chunkSizeOverride: 0, ownedHandle: null, ct);
        }

        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        [SuppressMessage("Performance", "CA1849:Call async methods when in an async method",
            Justification = "This factory is synchronous by design (see the VSTHRD200 suppression above); wrapping the caller-owned stream here, once, before any enumeration starts, is deliberate rather than a blocking call inside an async method.")]
        internal static IAsyncEnumerable<T> Create<T>(
            Stream stream, int degreeOfParallelism, CsvReaderOptions? readerOptions, ExcelParserConfig? config, CancellationToken ct)
        {
            CsvReaderOptions options = readerOptions ?? CsvReaderOptions.Default;
            ExcelParserConfig parserConfig = config ?? new ExcelParserConfig();
            int dop = Normalize(degreeOfParallelism);

            if (!CsvSourceResolver.TryResolve(stream, out CsvChunkSource source)
                || !CanPartition(dop, source.Length, options))
            {
                return Sequential<T>(Excel.FromCsv(stream, leaveOpen: true, options), parserConfig, ct);
            }
            return Build<T>(source, options, parserConfig, dop, chunkSizeOverride: 0, ownedHandle: null, ct);
        }

        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        internal static IAsyncEnumerable<T> CreateWithChunkSize<T>(
            ReadOnlyMemory<byte> data, int degreeOfParallelism, int chunkSize, CsvReaderOptions? readerOptions, ExcelParserConfig? config, CancellationToken ct)
        {
            CsvReaderOptions options = readerOptions ?? CsvReaderOptions.Default;
            ExcelParserConfig parserConfig = config ?? new ExcelParserConfig();
            return Build<T>(new CsvChunkSource(data), options, parserConfig, Normalize(degreeOfParallelism), chunkSize, ownedHandle: null, ct);
        }

        private static int Normalize(int degreeOfParallelism)
        {
            return degreeOfParallelism == 0 ? Environment.ProcessorCount : degreeOfParallelism;
        }

        private static bool CanPartition(int dop, long length, CsvReaderOptions options)
        {
            if (dop <= 1 || length < MinParallelBytes)
            {
                return false;
            }
            return options.Encoding is null || options.Encoding.CodePage == Encoding.UTF8.CodePage;
        }

        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        private static IAsyncEnumerable<T> Build<T>(
            CsvChunkSource source,
            CsvReaderOptions options,
            ExcelParserConfig config,
            int dop,
            int chunkSizeOverride,
            SafeFileHandle? ownedHandle,
            CancellationToken ct)
        {
            TypeMapInfo<T> info = TypeMapper<T>.GetCsvInfo();
            CsvBoundColumnMap<T> map;
            long dataStart;
            using (CsvReader headerReader = source.OpenReader(options))
            {
                map = CsvHeaderBinder.Bind<T>(headerReader, config, info, out dataStart);
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
                return Sequential<T>(source.OpenReader(options), config, ct);
            }

            return new ParallelCsvEnumerable<T>(source, plan, dataStart, map, info, options, config, dop, ownedHandle);
        }

        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        private static CsvEnumerable<T> Sequential<T>(
            CsvReader reader, ExcelParserConfig config, CancellationToken ct)
        {
            return new CsvEnumerable<T>(reader, config, ownsReader: true, ct);
        }

        private static async IAsyncEnumerable<T> Empty<T>()
        {
            await Task.CompletedTask.ConfigureAwait(false);
            yield break;
        }
    }
}
