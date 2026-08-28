using System.Buffers;
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
        //
        // Fills a caller-owned buffer rather than allocating one: the caller (CsvChunkWorker.
        // GuessStart) rents from ArrayPool and grows on demand, so a boundary scan costs no
        // allocation at all in the common case. Returns the number of bytes actually available.
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

        // The memory source's bytes are already in the process; boundary scanning can read them in
        // place instead of copying a window out. Only valid for IsMemory sources.
        internal ReadOnlySpan<byte> WindowSpan(long offset, int length)
        {
            long available = Length - offset;
            int want = (int)Math.Min(available, (long)length + 1);
            return want <= 0 ? default : _memory.Span.Slice((int)offset, want);
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
        // Boundary-scan window sizing. 4 KB holds any realistic CSV record several times over, and
        // stays far below the 85 KB Large Object Heap threshold so the rented buffer is an ordinary
        // Gen0 allocation the pool hands back immediately. Growth is geometric so even a
        // pathologically long record converges in a handful of reads.
        private const int InitialBoundaryWindow = 4 * 1024;
        private const int BoundaryWindowGrowth = 8;

        // No row-count estimation here any more. It existed to stop a chunk's model list from
        // doubling its way up to a Large Object Heap array, which mattered when a chunk was a share
        // of the file; with a 64 KB chunk and lists recycled across chunks by the merge (ListPool),
        // a list reaches its steady-state capacity within the first few chunks and never grows again.

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
            List<T> models,
            CancellationToken ct)
        {
            long start = confirmedStart ?? GuessStart(source, chunk, readerOptions.Quote);
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
            (long resolvedNextStart, ExcelParseException? failure, int failureRow) outcome;

            // The enumerator's stream constructor passes ownsSource: false (CsvReader.Enumerator.cs:65),
            // so it will NOT dispose this stream — the worker owns it. Disposing the RangedFileStream
            // is cheap and does not touch the shared SafeFileHandle, but leaving it undisposed trips
            // the IDisposable analyzers this repo builds with warnings-as-errors. The file/memory
            // branches are kept fully separate (rather than sharing one nullable partitionStream
            // variable) because that is the shape IDisposableAnalyzers can actually verify: each
            // resource's creation, its own `await using`, and its disposal all live in the same
            // lexical scope. Only the loop that *consumes* an already-constructed enumerator is
            // shared (ConsumeAsync below) — that part carries no disposal obligation of its own, so
            // moving it across a method boundary does not confuse the analyzer.
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

        // Drains an already-open, already-`await using`-guarded enumerator into `models`. Owns no
        // disposable of its own — `rows` is disposed by the caller's `await using` block — so it can
        // be shared between the memory and file branches without confusing IDisposableAnalyzers about
        // where the enumerator's lifetime ends.
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

        // Chunk 0 knows its start exactly (CsvHeaderBinder reported it). Every other chunk scans from
        // its nominal start under the Outside hypothesis. The scan is bounded by the chunk's own
        // length: a chunk with no boundary inside it holds no record start at all — its bytes belong
        // to a record an earlier chunk owns and overshoots into.
        //
        // The chunk's length bounds the scan, but it is NOT the window size. A record boundary sits
        // within the first few hundred bytes of a chunk in any realistic CSV, so the window starts
        // small and grows only against data pathological enough to need it. Reading the whole chunk
        // up front — which is what this did originally — allocated a chunk-sized array per chunk
        // (megabytes each, straight past the 85 KB Large Object Heap threshold) and read every byte
        // of the source a second time, to answer a question the first line almost always settles.
        // Measured on a 104 MB narrow-int corpus, that alone accounted for ~104 MB of LOH traffic
        // and every Gen2 collection the parallel path incurred over the sequential one.
        private static long GuessStart(CsvChunkSource source, CsvChunk chunk, byte quote)
        {
            long chunkLength = chunk.End - chunk.Start;
            if (chunkLength <= 0)
            {
                return long.MaxValue;
            }

            // Scanning in place, no window buffer at all: the bytes are already in the process.
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
                // +1 so a \r at the window's last position is resolved against the byte that
                // follows it rather than guessed at — the same reason the original read one extra
                // byte past the chunk.
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
                    // No boundary in this window. If it already covered the whole chunk, the chunk
                    // genuinely holds no record start; otherwise widen and rescan. Rescanning the
                    // prefix is deliberate over resuming: FindRecordStart is a left-to-right parity
                    // scan, so a wider window's answer is the same one a resumed scan would reach,
                    // and total rescan cost stays linear under geometric growth.
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
