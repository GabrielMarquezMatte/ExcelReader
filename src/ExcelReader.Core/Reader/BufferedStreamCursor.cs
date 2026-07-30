using System.Buffers;
using System.Runtime.InteropServices;

namespace ExcelReader.Core.Reader
{
    // Pooled-buffer refill/compact-or-grow cursor shared by the XLSX/XLSB/CSV forward-only stream
    // enumerators. Growth is capped the same way CellAccumulator caps its value buffer
    // (LimitChecks.NextBufferSize); each caller supplies its own (maxCellBytes, limitName) pair since
    // XLSX/XLSB draw the limit from ExcelReaderOptions while CSV draws it from CsvReaderOptions.
    internal sealed class BufferedStreamCursor
    {
        private const int InitialBuf = 64 * 1024;

        private readonly int _maxCellBytes;
        private readonly string _limitName;

        // False for a memory-backed cursor: Buf then aliases either the caller's own array (a ZIP
        // stored entry) or an array owned by ZipPart (a deflated entry) — never one this cursor rented,
        // so Return() must never hand it back to the pool.
        private readonly bool _pooled;

        internal byte[] Buf { get; private set; }
        internal int Pos { get; set; }
        internal int Len { get; private set; }
        internal bool Eof { get; private set; }

        internal BufferedStreamCursor(int maxCellBytes, string limitName, int initialCapacity = InitialBuf)
        {
            _maxCellBytes = maxCellBytes;
            _limitName = limitName;
            _pooled = true;
            Buf = ArrayPool<byte>.Shared.Rent(initialCapacity);
        }

        // Pre-filled, EOF from the start: the whole ZIP part is already decompressed, so there is
        // nothing left to refill and no source to refill from. `content` may alias a sub-range of a
        // larger array (e.g. a stored entry sliced out of the whole-file buffer), so Pos/Len start at
        // that sub-range's absolute offsets rather than always at 0 — every consumer of this cursor
        // (XlsxReader/XlsbReader included) indexes Buf directly via Pos/Len, so those offsets must stay
        // absolute instead of being rebased to the segment's own 0-based range.
        internal BufferedStreamCursor(ReadOnlyMemory<byte> content, int maxCellBytes, string limitName)
        {
            _maxCellBytes = maxCellBytes;
            _limitName = limitName;
            _pooled = false;
            Eof = true;
            if (MemoryMarshal.TryGetArray(content, out ArraySegment<byte> segment))
            {
                Buf = segment.Array!;
                Pos = segment.Offset;
                Len = segment.Offset + segment.Count;
                return;
            }
            // Rare: a non-array-backed ReadOnlyMemory<byte> (e.g. a custom MemoryManager<byte>). One
            // copy here is unavoidable since Buf must be a raw array.
            Buf = content.ToArray();
            Len = Buf.Length;
        }

        // Compact the consumed prefix, or grow the buffer if every byte is still unprocessed. Compacting
        // is skipped when the tail already has generous free room (>= a quarter of the buffer): the next
        // Fill can just read into that room directly, so moving the consumed prefix out of the way now
        // would be a memmove that pays for itself later than it needs to. Once the tail shrinks below
        // that threshold, compaction reclaims the (by-then-larger) consumed prefix in one pass, instead
        // of paying a smaller memmove on every single Fill call regardless of how much room already exists.
        private void PrepareBuffer()
        {
            int freeTail = Buf.Length - Len;
            if (freeTail > 0 && (Pos == 0 || freeTail >= Buf.Length / 4))
            {
                return;
            }
            if (Pos > 0)
            {
                Buf.AsSpan(Pos, Len - Pos).CopyTo(Buf);
                Len -= Pos;
                Pos = 0;
                return;
            }
            byte[] bigger = ArrayPool<byte>.Shared.Rent(LimitChecks.NextBufferSize(_maxCellBytes, _limitName, Buf.Length, Buf.Length + 1));
            Buf.AsSpan(0, Len).CopyTo(bigger);
            ArrayPool<byte>.Shared.Return(Buf);
            Buf = bigger;
        }

        // No-op once Eof is set at construction (the memory-backed ctor), rather than requiring every
        // enumerator call site to check Eof before calling Fill — there is no source to read from.
        internal void Fill(Stream? source)
        {
            if (source is null)
            {
                return;
            }
            PrepareBuffer();
            int n = source.Read(Buf, Len, Buf.Length - Len);
            if (n == 0)
            {
                Eof = true;
                return;
            }
            Len += n;
        }

        internal ValueTask FillAsync(Stream? source, CancellationToken ct)
        {
            return source is null ? ValueTask.CompletedTask : FillFromAsync(source, ct);
        }

        private async ValueTask FillFromAsync(Stream source, CancellationToken ct)
        {
            PrepareBuffer();
            int n = await source.ReadAsync(Buf.AsMemory(Len, Buf.Length - Len), ct).ConfigureAwait(false);
            if (n == 0)
            {
                Eof = true;
                return;
            }
            Len += n;
        }

        // Fast/slow split mirroring EnsureAsync below: the common case (already buffered) is a single
        // comparison the JIT can inline at call sites, instead of always entering a loop body.
        internal void Ensure(Stream? source, int n)
        {
            if (Len - Pos >= n || Eof)
            {
                return;
            }
            EnsureSlow(source, n);
        }

        private void EnsureSlow(Stream? source, int n)
        {
            while (Len - Pos < n && !Eof)
            {
                Fill(source);
            }
        }

        internal ValueTask EnsureAsync(Stream? source, int n, CancellationToken ct)
        {
            if (Len - Pos >= n || Eof)
            {
                return ValueTask.CompletedTask;
            }
            return EnsureSlowAsync(source, n, ct);
        }

        private async ValueTask EnsureSlowAsync(Stream? source, int n, CancellationToken ct)
        {
            while (Len - Pos < n && !Eof)
            {
                await FillAsync(source, ct).ConfigureAwait(false);
            }
        }

        internal void Return()
        {
            if (_pooled && Buf.Length > 0)
            {
                ArrayPool<byte>.Shared.Return(Buf);
            }
            Buf = [];
        }
    }
}
