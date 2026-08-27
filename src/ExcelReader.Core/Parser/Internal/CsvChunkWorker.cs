using ExcelReader.Core.Reader;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Core.Parser.Internal
{
    // One shape for both partitionable sources, so the worker has a single code path. A file is read
    // positionally off a shared handle; an in-memory source is simply sliced. Both extend to the end
    // of the source rather than to the chunk end, because a chunk must overshoot to finish the
    // record straddling its end.
    internal readonly struct CsvChunkSource
    {
        private readonly SafeFileHandle? _handle;
        private readonly ReadOnlyMemory<byte> _memory;
        private readonly long _startOffset;

        // startOffset lets a FileStream positioned mid-file be honored (Task 8's CsvSourceResolver
        // passes its Position). Every offset the rest of the pipeline works in stays relative to the
        // source's first byte; only this type knows about the file-absolute shift, which is what
        // keeps the chunk plan, worker, and merge free of the distinction.
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

        // A fresh sequential reader over this whole source. Used for the one-time header bind and for
        // the sequential fallback, and it must be repeatable: deriving it from the source rather than
        // from a caller-supplied Stream is what makes it so, since reading a Stream twice would start
        // the second read from the position the first one left behind.
        internal CsvReader OpenReader(CsvReaderOptions options)
        {
            if (IsMemory)
            {
                return Excel.FromCsv(SliceAt(0), options);
            }
            return Excel.FromCsv(OpenAt(0), leaveOpen: false, options);
        }

        // A window for boundary resolution: from `offset`, at most `length` bytes, plus one extra
        // byte. That extra byte is what makes a \r sitting at the chunk's last position resolvable
        // against real data instead of a guess about what follows it.
        internal byte[] ReadWindow(long offset, int length)
        {
            long available = Length - offset;
            int want = (int)Math.Min(available, (long)length + 1);
            if (want <= 0)
            {
                return [];
            }
            if (IsMemory)
            {
                return _memory.Slice((int)offset, want).ToArray();
            }
            byte[] buffer = new byte[want];
            int read = RandomAccess.Read(_handle!, buffer, _startOffset + offset);
            if (read == want)
            {
                return buffer;
            }
            Array.Resize(ref buffer, read);
            return buffer;
        }
    }

    // What one chunk contributes to the merged output, plus the two offsets the merge reconciles.
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

        // Where this chunk actually began parsing. A guess for every chunk but the first, checked
        // against the predecessor's ResolvedNextStart during the merge.
        internal long ActualStart { get; }

        // Offset of the first record starting at or after the chunk's nominal end — the ground truth
        // the *next* chunk's ActualStart is validated against. long.MaxValue when this chunk ran to
        // the end of the source.
        internal long ResolvedNextStart { get; }

        // A parse failure is carried, not thrown. Only the merge knows this chunk's global row
        // offset, so only the merge can raise it with the row number the sequential path would have
        // reported. FailureRowInChunk is zero-based within Models' record sequence.
        internal ExcelParseException? Failure { get; set; }

        internal int FailureRowInChunk { get; set; }
    }

    internal static class CsvChunkWorker
    {
        // Parses one chunk. `confirmedStart` is the offset the predecessor proved correct; when it is
        // null the worker guesses with the Outside hypothesis, which is right whenever no quoted
        // field straddles the chunk start — overwhelmingly the common case.
        internal static ValueTask<CsvChunkResult<T>> ParseAsync<T>(
            CsvChunkSource source,
            CsvChunk chunk,
            long? confirmedStart,
            CsvBoundColumnMap<T> map,
            TypeMapInfo<T> info,
            CsvReaderOptions readerOptions,
            ExcelParserConfig config,
            CancellationToken ct)
        {
            long start = confirmedStart ?? GuessStart(source, chunk, readerOptions.Quote);
            var models = new List<T>();
            if (start >= source.Length)
            {
                return new ValueTask<CsvChunkResult<T>>(new CsvChunkResult<T>(chunk.Index, models, start, long.MaxValue));
            }

            // A mid-file chunk must not strip a BOM: those three bytes are ordinary data there.
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
            long resolvedNextStart = long.MaxValue;
            ExcelParseException? failure = null;
            int failureRow = 0;

            // The enumerator's stream constructor passes ownsSource: false (CsvReader.Enumerator.cs:65),
            // so it will NOT dispose this stream — the worker owns it. Disposing the RangedFileStream
            // is cheap and does not touch the shared SafeFileHandle, but leaving it undisposed trips
            // the IDisposable analyzers this repo builds with warnings-as-errors. The file/memory
            // branches are kept fully separate (rather than sharing one nullable partitionStream
            // variable) because that is the shape IDisposableAnalyzers can actually verify: each
            // resource's creation and disposal live in the same lexical scope.
            if (source.IsMemory)
            {
                var rows = new CsvReader.Enumerator(source.SliceAt(start), options, ct);
                await using (rows.ConfigureAwait(false))
                {
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
                }
            }
            else
            {
                Stream partitionStream = source.OpenAt(start);
                try
                {
                    var rows = new CsvReader.Enumerator(partitionStream, options, ct);
                    await using (rows.ConfigureAwait(false))
                    {
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
                    }
                }
                finally
                {
                    await partitionStream.DisposeAsync().ConfigureAwait(false);
                }
            }

            return new CsvChunkResult<T>(chunk.Index, models, start, resolvedNextStart)
            {
                Failure = failure,
                FailureRowInChunk = failureRow,
            };
        }

        // Chunk 0 knows its start exactly (CsvHeaderBinder reported it). Every other chunk scans from
        // its nominal start under the Outside hypothesis. The window is the chunk's own length: a
        // chunk with no boundary inside it holds no record start at all — its bytes belong to a
        // record an earlier chunk owns and overshoots into.
        private static long GuessStart(CsvChunkSource source, CsvChunk chunk, byte quote)
        {
            int windowLength = (int)Math.Min(chunk.End - chunk.Start, int.MaxValue - 1);
            byte[] window = source.ReadWindow(chunk.Start, windowLength);
            int offset = CsvBoundaryResolver.FindRecordStart(window, quote, CsvQuoteParity.Outside);
            if (offset < 0)
            {
                return long.MaxValue;
            }
            return chunk.Start + offset;
        }
    }
}
