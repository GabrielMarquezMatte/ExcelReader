using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Reader;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Parser.Internal
{
    // Decides between the parallel pipeline and the ordinary sequential parser, and owns the shared
    // SafeFileHandle for the file-based path. Degradation is silent by design: both paths produce the
    // same sequence, so falling back can only cost speed, never correctness — and throwing instead
    // would make the API fragile in generic code that sometimes gets a file and sometimes does not.
    internal static class ParallelCsvFactory
    {
        // Below four cursor buffers there is nothing worth partitioning.
        private const long MinParallelBytes = 4 * 64 * 1024;

        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Factory method, not itself async; it synchronously decides a plan and returns a lazily-enumerated IAsyncEnumerable<T>, exactly like ExcelParser<T>.Parse(CsvReader).")]
        [SuppressMessage("Performance", "CA1849:Call async methods when in an async method",
            Justification = "This factory is synchronous by design (see the VSTHRD200 suppression above); opening the file here, once, before any enumeration starts, is deliberate rather than a blocking call inside an async method.")]
        [SuppressMessage("VisualStudio.Threading", "VSTHRD103:FromCsvFile synchronously blocks",
            Justification = "Same rationale as the CA1849 suppression: this factory method is synchronous by design.")]
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
        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Factory method, not itself async; it synchronously decides a plan and returns a lazily-enumerated IAsyncEnumerable<T>, exactly like ExcelParser<T>.Parse(CsvReader).")]
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
        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Factory method, not itself async; it synchronously decides a plan and returns a lazily-enumerated IAsyncEnumerable<T>, exactly like ExcelParser<T>.Parse(CsvReader).")]
        [SuppressMessage("Performance", "CA1849:Call async methods when in an async method",
            Justification = "This factory is synchronous by design (see the VSTHRD200 suppression above); wrapping the caller-owned stream here, once, before any enumeration starts, is deliberate rather than a blocking call inside an async method.")]
        [SuppressMessage("VisualStudio.Threading", "VSTHRD103:FromCsv synchronously blocks",
            Justification = "Same rationale as the CA1849 suppression: this factory method is synchronous by design.")]
        internal static IAsyncEnumerable<T> Create<T>(
            Stream stream, int degreeOfParallelism, CsvReaderOptions? readerOptions, ExcelParserConfig? config, CancellationToken ct)
        {
            CsvReaderOptions options = readerOptions ?? CsvReaderOptions.Default;
            ExcelParserConfig parserConfig = config ?? new ExcelParserConfig();
            int dop = Normalize(degreeOfParallelism);

            // leaveOpen: true throughout — the caller owns this stream on both paths.
            if (!CsvSourceResolver.TryResolve(stream, out CsvChunkSource source)
                || !CanPartition(dop, source.Length, options))
            {
                return Sequential<T>(Excel.FromCsv(stream, leaveOpen: true, options), parserConfig, ct);
            }
            // ownedHandle stays null even for a FileStream: that handle is borrowed from a stream the
            // caller still owns, so disposing it here would break them.
            return Build<T>(source, options, parserConfig, dop, chunkSizeOverride: 0, ownedHandle: null, ct);
        }

        // Test seam: forces a chunk size so a small fixture can be swept across every boundary
        // offset, and skips the minimum-size gate that would otherwise send every fixture down the
        // sequential path. Never reached by the public overloads.
        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Factory method, not itself async; it synchronously decides a plan and returns a lazily-enumerated IAsyncEnumerable<T>, exactly like ExcelParser<T>.Parse(CsvReader).")]
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
            // A non-UTF-8 source is transcoded through a sequential wrapper stream, which destroys the
            // byte-offset mapping partitioning depends on.
            return options.Encoding is null || options.Encoding.CodePage == Encoding.UTF8.CodePage;
        }

        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Factory helper, not itself async; it synchronously binds the header and plans chunks, then returns a lazily-enumerated IAsyncEnumerable<T>, exactly like ExcelParser<T>.Parse(CsvReader).")]
        // chunkSizeOverride sits next to dop (both shape the chunk plan) and before ct, which stays
        // last. It is not optional: a defaulted parameter here would have to trail the token.
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
            // Readers come from the source, never from a caller-supplied Stream: this runs twice on
            // the fallback path, and a Stream would resume the second read from wherever the first
            // one stopped.
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

            // The handle goes straight into the enumerable rather than into a wrapping
            // IAsyncEnumerable<T>. A wrapper would have to re-yield every row through a second async
            // iterator, and on a narrow-row corpus that per-row hop is a measurable share of the
            // single-threaded merge — the one part of this pipeline parallelism cannot shorten.
            return new ParallelCsvEnumerable<T>(source, plan, dataStart, map, info, options, config, dop, ownedHandle);
        }

        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        // The reader is handed to the enumerable to own, rather than wrapped in an `await using`
        // inside an async iterator. That wrapper is what this used to be, and it re-yielded every
        // single row through a second async state machine to buy one disposal at the end: measured on
        // an 8M-row narrow corpus, ~24 ns per row, or 190 ms of the 800 ms this path took — a quarter
        // of the sequential fallback's cost, spent on nothing.
        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Factory helper, not itself async: it constructs a lazily-enumerated IAsyncEnumerable<T> and returns it.")]
        private static CsvEnumerable<T> Sequential<T>(
            CsvReader reader, ExcelParserConfig config, CancellationToken ct)
        {
            return new CsvEnumerable<T>(reader, config, ownsReader: true, ct);
        }

        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "The enumerable also implements IAsyncEnumerable<T>; this is an async-iterator method, not an async entry point that awaits a single operation and returns its result.")]
        private static async IAsyncEnumerable<T> Empty<T>()
        {
            await Task.CompletedTask.ConfigureAwait(false);
            yield break;
        }
    }
}
