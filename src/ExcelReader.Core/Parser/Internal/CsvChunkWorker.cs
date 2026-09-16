using System.Buffers;
using ExcelReader.Core.Reader;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Parser.Internal
{
    internal readonly struct CsvChunkSource
    {
        private readonly SafeFileHandle? _handle;
        private readonly ReadOnlyMemory<byte> _memory;
        private readonly long _startOffset;

        internal CsvChunkSource(SafeFileHandle handle, long fileLength, long startOffset = 0)
        {
            _handle = handle;
            _memory = default;
            _startOffset = startOffset;
            Length = fileLength - startOffset;
        }

        internal CsvChunkSource(ReadOnlyMemory<byte> memory)
        {
            _handle = null;
            _memory = memory;
            _startOffset = 0;
            Length = memory.Length;
        }

        internal long Length { get; }

        internal bool IsMemory => _handle is null;

        internal Stream OpenAt(long offset)
        {
            return new RangedFileStream(_handle!, _startOffset + offset);
        }

        internal ReadOnlyMemory<byte> SliceAt(long offset)
        {
            return _memory[(int)offset..];
        }

        internal CsvReader OpenReader(CsvReaderOptions options)
        {
            if (IsMemory)
            {
                return Excel.FromCsv(SliceAt(0), options);
            }
            return Excel.FromCsv(OpenAt(0), leaveOpen: false, options);
        }

        internal int ReadWindow(long offset, Span<byte> buffer)
        {
            long available = Length - offset;
            int want = (int)Math.Min(available, buffer.Length);
            if (want <= 0)
            {
                return 0;
            }
            if (IsMemory)
            {
                _memory.Span.Slice((int)offset, want).CopyTo(buffer);
                return want;
            }
            return RandomAccess.Read(_handle!, buffer[..want], _startOffset + offset);
        }

        internal ReadOnlySpan<byte> WindowSpan(long offset, int length)
        {
            long available = Length - offset;
            int want = (int)Math.Min(available, (long)length + 1);
            return want <= 0 ? default : _memory.Span.Slice((int)offset, want);
        }
    }

    internal sealed class CsvChunkResult<T>
    {
        internal CsvChunkResult(int index, List<T> models, long actualStart, long resolvedNextStart)
        {
            Index = index;
            Models = models;
            ActualStart = actualStart;
            ResolvedNextStart = resolvedNextStart;
        }

        internal int Index { get; }

        internal List<T> Models { get; }

        internal long ActualStart { get; }

        internal long ResolvedNextStart { get; }

        internal ExcelParseException? Failure { get; set; }

        internal int FailureRowInChunk { get; set; }
    }

    internal static class CsvChunkWorker
    {
        private const int InitialBoundaryWindow = 4 * 1024;
        private const int BoundaryWindowGrowth = 8;

        internal static ValueTask<CsvChunkResult<T>> ParseAsync<T>(
            CsvChunkSource source,
            CsvChunk chunk,
            long? confirmedStart,
            CsvBoundColumnMap<T> map,
            TypeMapInfo<T> info,
            CsvReaderOptions readerOptions,
            ExcelParserConfig config,
            List<T> models,
            CancellationToken ct)
        {
            long start = confirmedStart ?? GuessStart(source, chunk, readerOptions.Quote);
            if (start >= source.Length)
            {
                return new ValueTask<CsvChunkResult<T>>(new CsvChunkResult<T>(chunk.Index, models, start, long.MaxValue));
            }

            CsvReaderOptions chunkOptions = readerOptions with { DetectEncodingFromByteOrderMark = start == 0 };
            return ParseFromAsync(source, chunk, start, models, map, info, chunkOptions, config, ct);
        }

        private static async ValueTask<CsvChunkResult<T>> ParseFromAsync<T>(
            CsvChunkSource source,
            CsvChunk chunk,
            long start,
            List<T> models,
            CsvBoundColumnMap<T> map,
            TypeMapInfo<T> info,
            CsvReaderOptions options,
            ExcelParserConfig config,
            CancellationToken ct)
        {
            var projector = new CsvRowProjector<T>(info, map, config.Culture, config.ThrowOnParseFailure);
            (long resolvedNextStart, ExcelParseException? failure, int failureRow) outcome;

            if (source.IsMemory)
            {
                var rows = new CsvReader.Enumerator(source.SliceAt(start), options, ct);
                await using (rows.ConfigureAwait(false))
                {
                    outcome = await ConsumeAsync(rows, start, chunk, projector, models).ConfigureAwait(false);
                }
                return new CsvChunkResult<T>(chunk.Index, models, start, outcome.resolvedNextStart)
                {
                    Failure = outcome.failure,
                    FailureRowInChunk = outcome.failureRow,
                };
            }
            Stream partitionStream = source.OpenAt(start);
            try
            {
                var outerRows = new CsvReader.Enumerator(partitionStream, options, ct);
                await using (outerRows.ConfigureAwait(false))
                {
                    outcome = await ConsumeAsync(outerRows, start, chunk, projector, models).ConfigureAwait(false);
                }
            }
            finally
            {
                await partitionStream.DisposeAsync().ConfigureAwait(false);
            }
            return new CsvChunkResult<T>(chunk.Index, models, start, outcome.resolvedNextStart)
            {
                Failure = outcome.failure,
                FailureRowInChunk = outcome.failureRow,
            };
        }

        private static async ValueTask<(long resolvedNextStart, ExcelParseException? failure, int failureRow)> ConsumeAsync<T>(
            CsvReader.Enumerator rows,
            long start,
            CsvChunk chunk,
            CsvRowProjector<T> projector,
            List<T> models)
        {
            long resolvedNextStart = long.MaxValue;
            ExcelParseException? failure = null;
            int failureRow = 0;
            while (await rows.MoveNextAsync().ConfigureAwait(false))
            {
                long absolute = start + rows.CurrentRecordStart;
                if (absolute >= chunk.End)
                {
                    resolvedNextStart = absolute;
                    break;
                }

                T model = default!;
                try
                {
                    if (projector.Advance(rows, ref model) == ProjectionStep.Yield)
                    {
                        models.Add(model);
                    }
                }
                catch (ExcelParseException ex)
                {
                    failure = ex;
                    failureRow = models.Count;
                    break;
                }
            }
            return (resolvedNextStart, failure, failureRow);
        }

        private static long GuessStart(CsvChunkSource source, CsvChunk chunk, byte quote)
        {
            long chunkLength = chunk.End - chunk.Start;
            if (chunkLength <= 0)
            {
                return long.MaxValue;
            }

            if (source.IsMemory)
            {
                int cap = (int)Math.Min(chunkLength, int.MaxValue - 1);
                int found = CsvBoundaryResolver.FindRecordStart(
                    source.WindowSpan(chunk.Start, cap), quote, CsvQuoteParity.Outside);
                return found < 0 ? long.MaxValue : chunk.Start + found;
            }

            long remaining = Math.Min(chunkLength, source.Length - chunk.Start);
            int windowLength = (int)Math.Min(remaining, InitialBoundaryWindow);
            while (true)
            {
                int rentSize = windowLength == remaining ? windowLength + 1 : windowLength;
                byte[] buffer = ArrayPool<byte>.Shared.Rent(rentSize);
                try
                {
                    int read = source.ReadWindow(chunk.Start, buffer.AsSpan(0, rentSize));
                    if (read <= 0)
                    {
                        return long.MaxValue;
                    }
                    int found = CsvBoundaryResolver.FindRecordStart(
                        buffer.AsSpan(0, read), quote, CsvQuoteParity.Outside);
                    if (found >= 0)
                    {
                        return chunk.Start + found;
                    }
                    if (read >= remaining || windowLength >= remaining)
                    {
                        return long.MaxValue;
                    }
                }
                finally
                {
                    ArrayPool<byte>.Shared.Return(buffer);
                }
                windowLength = (int)Math.Min(remaining, (long)windowLength * BoundaryWindowGrowth);
            }
        }
    }
}
