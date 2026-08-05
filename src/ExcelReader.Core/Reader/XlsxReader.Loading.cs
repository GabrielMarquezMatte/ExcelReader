using System.Buffers;
using System.IO.Compression;

namespace ExcelReader.Core.Reader
{
    public sealed partial class XlsxReader
    {
        // --- workbook / shared-strings loading (one-time, small except sharedStrings) ---
        private static (string Name, string Path)[] ParseSheets(ReadOnlySpan<byte> wbBytes, ReadOnlySpan<byte> relsBytes)
        {
            if (wbBytes.IsEmpty)
            {
                return [];
            }
            Dictionary<string, string> rels = XlsxXml.ParseRelationships(relsBytes);
            var sheets = new List<(string, string)>();
            // Some producers prefix every element (<x:workbook>/<x:sheet>); match the prefixed name
            // when present so a prefixed workbook part still yields its sheets instead of none.
            ReadOnlySpan<byte> prefix = XlsxXml.DetectElementPrefix(wbBytes);
            ReadOnlySpan<byte> sheetTag = "<sheet "u8;
            if (!prefix.IsEmpty)
            {
                sheetTag = XlsxXml.Token("<"u8, prefix, "sheet "u8);
            }
            foreach (var tag in Tags(wbBytes, sheetTag))
            {
                var rid = XlsxXml.DecodeToString(XlsxXml.Attr(tag, " r:id="u8));
                if (rels.TryGetValue(rid, out var target))
                {
                    var name = XlsxXml.DecodeToString(XlsxXml.Attr(tag, " name="u8));
                    sheets.Add((name, XlsxXml.NormalizePart(target)));
                }
            }
            return [.. sheets];
        }

        // Returns true when xl/workbook.xml contains <workbookPr date1904="1"> (the Mac epoch).
        private static bool ParseDate1904(ReadOnlySpan<byte> src)
        {
            if (src.IsEmpty)
            {
                return false;
            }
            ReadOnlySpan<byte> prefix = XlsxXml.DetectElementPrefix(src);
            ReadOnlySpan<byte> workbookPrTag = "<workbookPr"u8;
            if (!prefix.IsEmpty)
            {
                workbookPrTag = XlsxXml.Token("<"u8, prefix, "workbookPr"u8);
            }
            int pos = IdxOf(src, 0, workbookPrTag);
            if (pos < 0)
            {
                return false;
            }
            int end = IdxOf(src, pos, (byte)'>');
            if (end < 0)
            {
                return false;
            }
            var attr = XlsxXml.Attr(src.Slice(pos, end - pos + 1), " date1904="u8);
            return attr.SequenceEqual("1"u8) || attr.SequenceEqual("true"u8);
        }

        private void EnsureSharedLoaded()
        {
            if (_sharedLoaded)
            {
                return;
            }
            _sharedLoaded = true;
            ZipArchiveEntry? entry = _zip!.GetEntry("xl/sharedStrings.xml");
            if (entry is null)
            {
                return;
            }
            WorkbookLookups.ThrowIfSharedEntryTooLarge(entry.Length, _decompressedBytes, _options);
            using LimitedReadStream stream = WorkbookLookups.OpenEntryStream(entry, _decompressedBytes, _options,
                nameof(ExcelReaderOptions.MaxSharedStringBytes), _options.MaxSharedStringBytes);
            ParseSharedStreaming(stream, entry.Length);
        }

        private async ValueTask EnsureSharedLoadedAsync(CancellationToken ct)
        {
            if (_sharedLoaded)
            {
                return;
            }
            _sharedLoaded = true;
            ZipArchiveEntry? entry = _zip!.GetEntry("xl/sharedStrings.xml");
            if (entry is null)
            {
                return;
            }
            WorkbookLookups.ThrowIfSharedEntryTooLarge(entry.Length, _decompressedBytes, _options);
            LimitedReadStream stream = await WorkbookLookups.OpenEntryStreamAsync(
                entry, _decompressedBytes, _options, ct,
                nameof(ExcelReaderOptions.MaxSharedStringBytes), _options.MaxSharedStringBytes).ConfigureAwait(false);
            await using (stream.ConfigureAwait(false))
            {
                await ParseSharedStreamingAsync(stream, entry.Length, ct).ConfigureAwait(false);
            }
        }

        // Both the sync and async shared-strings parsers open on the same footing: reject a declared
        // length that cannot even be indexed, size the flat buffer from it (decoded text is never
        // longer than its XML source), and build a cursor whose growth is capped by MaxSharedStringBytes.
        // The flat-buffer Rent stays in each caller's own try (not here) so a throw from this method
        // itself never leaves a rented buffer behind for the caller's finally to miss.
        private BufferedStreamCursor CreateSharedCursor(long entryLength, out int partLength)
        {
            LimitChecks.ThrowIfEntryLengthExceeds(entryLength, Array.MaxLength, "ArrayMaxLength");
            partLength = (int)entryLength;
            return new BufferedStreamCursor(SharedFlatGrowthCap(), nameof(ExcelReaderOptions.MaxSharedStringBytes),
                WorkbookLookups.InitialBufferCapacity(entryLength));
        }

        // Streams xl/sharedStrings.xml through a growable pooled buffer instead of inflating the whole
        // part before parsing a byte of it — mirrors how the row enumerators use BufferedStreamCursor/
        // EnsureRowBuffered so decompression overlaps the scan (via PrefetchStream) instead of finishing
        // first. Growth is capped by MaxSharedStringBytes, the same limit ThrowIfSharedEntryTooLarge
        // already checked the declared part length against.
        private void ParseSharedStreaming(Stream stream, long entryLength)
        {
            BufferedStreamCursor io = CreateSharedCursor(entryLength, out int partLength);
            try
            {
                // Decoded text is never longer than its XML, so partLength bounds the flat buffer —
                // identical sizing to the inflate-all-then-parse path this replaces.
                _sharedFlat = ArrayPool<byte>.Shared.Rent(Math.Max(1, partLength));
                _sharedOffsets = ParseSharedBody(io, stream, partLength);
            }
            finally
            {
                io.Return();
            }
        }

        private async ValueTask ParseSharedStreamingAsync(Stream stream, long entryLength, CancellationToken ct)
        {
            BufferedStreamCursor io = CreateSharedCursor(entryLength, out int partLength);
            try
            {
                _sharedFlat = ArrayPool<byte>.Shared.Rent(Math.Max(1, partLength));
                _sharedOffsets = await ParseSharedBodyAsync(io, stream, partLength, ct).ConfigureAwait(false);
            }
            finally
            {
                io.Return();
            }
        }

        // Parses the <sst>/<si>/<t> structure one growable-buffer window at a time. `io.Pos` doubles
        // as the "everything before this is fully consumed" marker BufferedStreamCursor.Fill uses to
        // compact — every search below runs from io.Pos and the caller advances it as soon as bytes
        // before the new position are no longer needed, so a Fill mid-search never invalidates an
        // already-found offset (see FindSeqGrowing/EnsureSiBuffered).
        private int[] ParseSharedBody(BufferedStreamCursor io, Stream? stream, int partLength)
        {
            io.Ensure(stream, 256); // root element + its xmlns declarations sit at the head of the part
            var tok = new SharedStringTokens(XlsxXml.DetectElementPrefix(io.Buf.AsSpan(0, io.Len)));

            int uniqueCount = 0;
            int sstPos = FindSeqGrowing(io, stream, tok.SstTag);
            if (sstPos >= 0)
            {
                io.Pos = sstPos;
                int sstEnd = FindSeqGrowing(io, stream, GtToken);
                if (sstEnd > sstPos)
                {
                    uniqueCount = XlsxXml.ParseIntOr(XlsxXml.Attr(io.Buf.AsSpan(sstPos, sstEnd - sstPos), " uniqueCount="u8), 0);
                }
            }
            LimitChecks.ThrowIfSharedStringCountImplausible(uniqueCount, partLength);

            int[] offsets = new int[uniqueCount > 0 ? uniqueCount + 1 : Math.Max(16, partLength / 64)];
            int offsetCount = 1;
            int flat = 0;
            while (true)
            {
                int si = FindSeqGrowing(io, stream, tok.SiTag);
                if (si < 0)
                {
                    break;
                }
                io.Pos = si;
                (int open, int close) = EnsureSiBuffered(io, stream, tok.SiClose);
                if (open < 0)
                {
                    break;
                }
                flat = AppendSharedEntry(io, tok, open, close, flat, out int nextPos);
                AddSharedOffset(ref offsets, ref offsetCount, flat);
                io.Pos = nextPos;
            }
            if (offsetCount != offsets.Length)
            {
                Array.Resize(ref offsets, offsetCount);
            }
            return offsets;
        }

        private async ValueTask<int[]> ParseSharedBodyAsync(BufferedStreamCursor io, Stream stream, int partLength, CancellationToken ct)
        {
            await io.EnsureAsync(stream, 256, ct).ConfigureAwait(false);
            var tok = new SharedStringTokens(XlsxXml.DetectElementPrefix(io.Buf.AsSpan(0, io.Len)));

            int uniqueCount = 0;
            int sstPos = await FindSeqGrowingAsync(io, stream, tok.SstTag, ct).ConfigureAwait(false);
            if (sstPos >= 0)
            {
                io.Pos = sstPos;
                int sstEnd = await FindSeqGrowingAsync(io, stream, GtToken, ct).ConfigureAwait(false);
                if (sstEnd > sstPos)
                {
                    uniqueCount = XlsxXml.ParseIntOr(XlsxXml.Attr(io.Buf.AsSpan(sstPos, sstEnd - sstPos), " uniqueCount="u8), 0);
                }
            }
            LimitChecks.ThrowIfSharedStringCountImplausible(uniqueCount, partLength);

            int[] offsets = new int[uniqueCount > 0 ? uniqueCount + 1 : Math.Max(16, partLength / 64)];
            int offsetCount = 1;
            int flat = 0;
            while (true)
            {
                int si = await FindSeqGrowingAsync(io, stream, tok.SiTag, ct).ConfigureAwait(false);
                if (si < 0)
                {
                    break;
                }
                io.Pos = si;
                (int open, int close) = await EnsureSiBufferedAsync(io, stream, tok.SiClose, ct).ConfigureAwait(false);
                if (open < 0)
                {
                    break;
                }
                flat = AppendSharedEntry(io, tok, open, close, flat, out int nextPos);
                AddSharedOffset(ref offsets, ref offsetCount, flat);
                io.Pos = nextPos;
            }
            if (offsetCount != offsets.Length)
            {
                Array.Resize(ref offsets, offsetCount);
            }
            return offsets;
        }

        // Decodes the <si> element starting at `open` (the '>' ending its open tag — self-closing or
        // not, already guaranteed fully buffered by EnsureSiBuffered(Async)) into _sharedFlat, and
        // reports where the next element search should resume via `nextPos`. `close` is the absolute
        // position of "</si>"'s own '<' as already located by EnsureSiBuffered(Async) — passed through
        // instead of re-running the same IndexOf(siClose) scan a second time (-1 for a self-closing
        // <si/>, which has no body to locate). No Fill/FillAsync happens here, so plain bounded index
        // math is safe exactly like ParseRow's post-EnsureRowBuffered body.
        private int AppendSharedEntry(BufferedStreamCursor io, SharedStringTokens tok, int open, int close, int flat, out int nextPos)
        {
            if (close < 0) // <si/>: no body
            {
                nextPos = open + 1;
                return flat;
            }
            int inner = close - open - 1;
            EnsureSharedFlat(flat + inner, flat);
            int written = XlsxXml.WriteTextRuns(io.Buf.AsSpan(open + 1, inner), _sharedFlat.AsSpan(flat),
                tok.TOpen, tok.TClose, tok.RPhOpen, tok.RPhClose);
            nextPos = close + tok.SiClose.Length;
            return flat + written;
        }

        // The declared central-directory length bounds the flat buffer only for a well-formed part.
        // The streaming reader consumes to the entry's real EOF rather than stopping at that declared
        // length (which ReadExactly used to enforce), so an entry that under-reports it would otherwise
        // run past the buffer and surface a raw index exception. Grow on what is actually decoded and
        // let MaxSharedStringBytes be what stops it, so the failure stays ExcelLimitExceededException.
        private void EnsureSharedFlat(int needed, int live)
        {
            if (needed <= _sharedFlat.Length)
            {
                return;
            }
            byte[] bigger = ArrayPool<byte>.Shared.Rent(LimitChecks.NextBufferSize(
                SharedFlatGrowthCap(), nameof(ExcelReaderOptions.MaxSharedStringBytes), _sharedFlat.Length, needed));
            Array.Copy(_sharedFlat, bigger, live);
            ArrayPool<byte>.Shared.Return(_sharedFlat);
            _sharedFlat = bigger;
        }

        private int SharedFlatGrowthCap()
        {
            if (_options.MaxSharedStringBytes <= 0)
            {
                return 0;
            }
            return (int)Math.Min(_options.MaxSharedStringBytes, Array.MaxLength);
        }

        // Single-byte '>' searches reuse the sequence search below; IndexOf handles a length-1
        // needle, so a dedicated overload would only duplicate the grow loop.
        private static readonly byte[] GtToken = ">"u8.ToArray();

        // Looks for `seq` in the currently buffered window [io.Pos..io.Len). -1 means "not in this
        // window" — the caller decides whether that is EOF (give up) or a reason to Fill and retry.
        // Split out so the sync and async growth loops share one search instead of two copies of the
        // same IndexOf.
        private static int FindSeqInWindow(BufferedStreamCursor io, byte[] seq)
        {
            int rel = io.Buf.AsSpan(io.Pos, io.Len - io.Pos).IndexOf(seq);
            return rel < 0 ? -1 : io.Pos + rel;
        }

        // Grows io (via Fill) until `seq` is found at or after io.Pos, or the stream ends. The caller
        // owns io.Pos as the search anchor — set it immediately before calling whenever the anchor
        // should move, so a Fill-triggered compaction (which always resets io.Pos to 0) never strands a
        // position computed against the pre-compaction buffer layout.
        private static int FindSeqGrowing(BufferedStreamCursor io, Stream? stream, byte[] seq)
        {
            while (true)
            {
                int found = FindSeqInWindow(io, seq);
                if (found >= 0)
                {
                    return found;
                }
                if (io.Eof)
                {
                    return -1;
                }
                io.Fill(stream);
            }
        }

        private static async ValueTask<int> FindSeqGrowingAsync(BufferedStreamCursor io, Stream stream, byte[] seq, CancellationToken ct)
        {
            while (true)
            {
                int found = FindSeqInWindow(io, seq);
                if (found >= 0)
                {
                    return found;
                }
                if (io.Eof)
                {
                    return -1;
                }
                await io.FillAsync(stream, ct).ConfigureAwait(false);
            }
        }

        // The '>' that closes the <si ...> open tag, but only once the whole element is contiguous in
        // the buffer: either the tag self-closes (<si/>) or its matching "</si>" is already buffered
        // too. Open is -1 when "keep filling" — same contract as FindSeqInWindow. Close carries the
        // absolute position of "</si>"'s own '<' (found by this same scan) so AppendSharedEntry never
        // has to re-run the identical IndexOf(siClose) search a second time; -1 for a self-closing
        // <si/>, which has no body to locate.
        private static (int Open, int Close) FindSiEndInWindow(BufferedStreamCursor io, byte[] siClose)
        {
            int openRel = io.Buf.AsSpan(io.Pos, io.Len - io.Pos).IndexOf((byte)'>');
            if (openRel < 0)
            {
                return (-1, -1);
            }
            int open = io.Pos + openRel;
            if (io.Buf[open - 1] == (byte)'/')
            {
                return (open, -1); // self-closing, no body
            }
            int rel = io.Buf.AsSpan(open, io.Len - open).IndexOf(siClose);
            if (rel < 0)
            {
                return (-1, -1);
            }
            return (open, open + rel);
        }

        // Grows io (io.Pos already anchored at the '<si' tag's start by the caller) until the whole
        // element — open tag through "</si>", or through a self-closing "<si .../>"'s own '>' — sits
        // contiguously in io.Buf. Mirrors XlsxReader.Enumerator.EnsureRowBuffered's "buffer the whole
        // element before parsing it" contract; returns -1 on a truncated file, matching the original
        // ParseShared's own break-on-truncation behavior instead of throwing.
        private static (int Open, int Close) EnsureSiBuffered(BufferedStreamCursor io, Stream? stream, byte[] siClose)
        {
            while (true)
            {
                (int open, int close) = FindSiEndInWindow(io, siClose);
                if (open >= 0)
                {
                    return (open, close);
                }
                if (io.Eof)
                {
                    return (-1, -1);
                }
                io.Fill(stream);
            }
        }

        private static async ValueTask<(int Open, int Close)> EnsureSiBufferedAsync(BufferedStreamCursor io, Stream stream, byte[] siClose, CancellationToken ct)
        {
            while (true)
            {
                (int open, int close) = FindSiEndInWindow(io, siClose);
                if (open >= 0)
                {
                    return (open, close);
                }
                if (io.Eof)
                {
                    return (-1, -1);
                }
                await io.FillAsync(stream, ct).ConfigureAwait(false);
            }
        }

        private static void AddSharedOffset(ref int[] offsets, ref int count, int value)
        {
            if (count == offsets.Length)
            {
                Array.Resize(ref offsets, offsets.Length * 2);
            }
            offsets[count++] = value;
        }

        private static int IdxOf(ReadOnlySpan<byte> s, int from, ReadOnlySpan<byte> seq)
        {
            int r = s[from..].IndexOf(seq);
            return r < 0 ? -1 : r + from;
        }

        private static int IdxOf(ReadOnlySpan<byte> s, int from, byte b)
        {
            int r = s[from..].IndexOf(b);
            return r < 0 ? -1 : r + from;
        }

        private static TagSpanEnumerable Tags(ReadOnlySpan<byte> buf, ReadOnlySpan<byte> prefix)
        {
            return new TagSpanEnumerable(buf, prefix);
        }
    }
}
