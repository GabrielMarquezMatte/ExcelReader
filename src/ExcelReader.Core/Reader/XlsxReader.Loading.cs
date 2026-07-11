using System.Buffers;

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
            foreach (var tag in Tags(wbBytes, "<sheet "u8))
            {
                var rid = XlsxXml.DecodeToString(XlsxXml.Attr(tag, " r:id=\""u8));
                if (rels.TryGetValue(rid, out var target))
                {
                    var name = XlsxXml.DecodeToString(XlsxXml.Attr(tag, " name=\""u8));
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
            int pos = IdxOf(src, 0, "<workbookPr"u8);
            if (pos < 0)
            {
                return false;
            }
            int end = IdxOf(src, pos, (byte)'>');
            if (end < 0)
            {
                return false;
            }
            var attr = XlsxXml.Attr(src.Slice(pos, end - pos + 1), " date1904=\""u8);
            return attr.SequenceEqual("1"u8) || attr.SequenceEqual("true"u8);
        }

        private void EnsureSharedLoaded()
        {
            if (_sharedLoaded)
            {
                return;
            }
            _sharedLoaded = true;
            var entry = _zip.GetEntry("xl/sharedStrings.xml");
            if (entry is not null)
            {
                var (bytes, length) = ZipEntryBytes.ReadPooled(
                    entry,
                    _decompressedBytes,
                    nameof(ExcelReaderOptions.MaxSharedStringBytes),
                    _options.MaxSharedStringBytes);
                try
                {
                    ParseShared(bytes.AsSpan(0, length));
                }
                finally
                {
                    ArrayPool<byte>.Shared.Return(bytes);
                }
            }
        }

        private async ValueTask EnsureSharedLoadedAsync(CancellationToken ct)
        {
            if (_sharedLoaded)
            {
                return;
            }
            _sharedLoaded = true;
            var entry = _zip.GetEntry("xl/sharedStrings.xml");
            if (entry is not null)
            {
                var (bytes, length) = await ZipEntryBytes.ReadPooledAsync(
                    entry,
                    _decompressedBytes,
                    ct,
                    nameof(ExcelReaderOptions.MaxSharedStringBytes),
                    _options.MaxSharedStringBytes).ConfigureAwait(false);
                try
                {
                    ParseShared(bytes.AsSpan(0, length));
                }
                finally
                {
                    ArrayPool<byte>.Shared.Return(bytes);
                }
            }
        }

        private void ParseShared(ReadOnlySpan<byte> src)
        {
            LimitChecks.ThrowIfOverSharedStringLimit(_options, src.Length);
            // Decoded text is never longer than its XML, so src.Length bounds the flat buffer.
            _sharedFlat = ArrayPool<byte>.Shared.Rent(Math.Max(1, src.Length));

            // Pre-size offsets from <sst uniqueCount="N">; exact counts avoid a final array copy.
            int uniqueCount = 0;
            int sstPos = IdxOf(src, 0, "<sst"u8);
            if (sstPos >= 0)
            {
                int sstEnd = IdxOf(src, sstPos, (byte)'>');
                if (sstEnd > sstPos)
                {
                    uniqueCount = ParseIntOr(XlsxXml.Attr(src[sstPos..sstEnd], " uniqueCount=\""u8), 0);
                }
            }

            // When uniqueCount is absent, guess from the XML size instead of always starting at 16 —
            // avoids repeated Array.Resize doublings (each a full copy) on large string-heavy sheets.
            int[] offsets = new int[uniqueCount > 0 ? uniqueCount + 1 : Math.Max(16, src.Length / 64)];
            int offsetCount = 1;
            int flat = 0;
            int p = 0;
            while (true)
            {
                int si = IdxOf(src, p, "<si"u8);
                if (si < 0)
                {
                    break;
                }
                int open = IdxOf(src, si, (byte)'>');
                if (open < 0)
                {
                    break;
                }
                if (src[open - 1] != '/') // not <si/>
                {
                    int end = IdxOf(src, open, "</si>"u8);
                    if (end < 0)
                    {
                        break;
                    }
                    flat += XlsxXml.WriteTextRuns(src.Slice(open + 1, end - open - 1), _sharedFlat.AsSpan(flat));
                    p = end + 5;
                }
                else
                {
                    p = open + 1;
                }
                AddSharedOffset(ref offsets, ref offsetCount, flat);
            }
            if (offsetCount != offsets.Length)
            {
                Array.Resize(ref offsets, offsetCount);
            }
            _sharedOffsets = offsets;
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
