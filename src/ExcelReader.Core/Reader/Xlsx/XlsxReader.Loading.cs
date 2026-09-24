using System.Buffers;
using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Zip;

namespace ExcelReader.Core.Reader.Xlsx
{
    public sealed partial class XlsxReader
    {
        private static (string Name, string Path, ExcelSheetVisibility Visibility)[] ParseSheets(ReadOnlySpan<byte> wbBytes, ReadOnlySpan<byte> relsBytes)
        {
            if (wbBytes.IsEmpty)
            {
                return [];
            }
            Dictionary<string, string> rels = XlsxXml.ParseRelationships(relsBytes);
            var sheets = new List<(string, string, ExcelSheetVisibility)>();
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
                    sheets.Add((name, XlsxXml.NormalizePart(target), ParseVisibility(XlsxXml.Attr(tag, " state="u8))));
                }
            }
            return [.. sheets];
        }

        private static ExcelSheetVisibility ParseVisibility(ReadOnlySpan<byte> state)
        {
            if (Ascii.EqualsIgnoreCase(state, "hidden"u8))
            {
                return ExcelSheetVisibility.Hidden;
            }
            return Ascii.EqualsIgnoreCase(state, "veryHidden"u8) ? ExcelSheetVisibility.VeryHidden : ExcelSheetVisibility.Visible;
        }

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

        private BufferedStreamCursor CreateSharedCursor(long entryLength, out int partLength)
        {
            LimitChecks.ThrowIfEntryLengthExceeds(entryLength, Array.MaxLength, "ArrayMaxLength");
            partLength = (int)entryLength;
            return new BufferedStreamCursor(SharedFlatGrowthCap(), nameof(ExcelReaderOptions.MaxSharedStringBytes),
                WorkbookLookups.InitialBufferCapacity(entryLength));
        }

        private void ParseSharedStreaming(Stream stream, long entryLength)
        {
            BufferedStreamCursor io = CreateSharedCursor(entryLength, out int partLength);
            try
            {
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

        private int[] ParseSharedBody(BufferedStreamCursor io, Stream? stream, int partLength)
        {
            io.Ensure(stream, 256);
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

        private int AppendSharedEntry(BufferedStreamCursor io, SharedStringTokens tok, int open, int close, int flat, out int nextPos)
        {
            if (close < 0)
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

        private static readonly byte[] GtToken = ">"u8.ToArray();

        private static int FindSeqInWindow(BufferedStreamCursor io, byte[] seq)
        {
            int rel = io.Buf.AsSpan(io.Pos, io.Len - io.Pos).IndexOf(seq);
            return rel < 0 ? -1 : io.Pos + rel;
        }

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
                return (open, -1);
            }
            int rel = io.Buf.AsSpan(open, io.Len - open).IndexOf(siClose);
            if (rel < 0)
            {
                return (-1, -1);
            }
            return (open, open + rel);
        }

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
            if (from < 0)
            {
                return -1;
            }
            int r = s[from..].IndexOf(seq);
            return r < 0 ? -1 : r + from;
        }

        private static int IdxOf(ReadOnlySpan<byte> s, int from, byte b)
        {
            if (from < 0)
            {
                return -1;
            }
            int r = s[from..].IndexOf(b);
            return r < 0 ? -1 : r + from;
        }

        private static TagSpanEnumerable Tags(ReadOnlySpan<byte> buf, ReadOnlySpan<byte> prefix)
        {
            return new TagSpanEnumerable(buf, prefix);
        }
    }
}
