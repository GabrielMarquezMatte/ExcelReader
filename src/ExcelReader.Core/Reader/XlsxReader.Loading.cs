using System.Buffers;
using System.IO.Compression;

namespace ExcelReader.Core.Reader
{
    public sealed partial class XlsxReader
    {
        // --- workbook / shared-strings loading (one-time, small except sharedStrings) ---

        private static readonly byte[] _relationshipTag = "<Relationship"u8.ToArray();
        private static readonly byte[] _sheetTag = "<sheet "u8.ToArray();

        private static (string Name, string Path)[] ParseSheets(ReadOnlyMemory<byte> wbBytes, ReadOnlyMemory<byte> relsBytes)
        {
            if (wbBytes.IsEmpty)
            {
                return [];
            }
            // rId -> target part path
            Dictionary<string, string> rels = new(StringComparer.Ordinal);
            if (!relsBytes.IsEmpty)
            {
                foreach (var tag in Tags(relsBytes, _relationshipTag))
                {
                    var id = Decode(XlsxXml.Attr(tag.Span, " Id=\""u8));
                    var target = Decode(XlsxXml.Attr(tag.Span, " Target=\""u8));
                    if (id.Length > 0)
                    {
                        rels[id] = target;
                    }
                }
            }
            var sheets = new List<(string, string)>();
            foreach (var tag in Tags(wbBytes, _sheetTag))
            {
                var name = Decode(XlsxXml.Attr(tag.Span, " name=\""u8));
                var rid = Decode(XlsxXml.Attr(tag.Span, " r:id=\""u8));
                if (rels.TryGetValue(rid, out var target))
                {
                    sheets.Add((name, NormalizePart(target)));
                }
            }
            return [.. sheets];
        }

        private static string NormalizePart(ReadOnlySpan<char> target)
        {
            if (target.Length > 0 && target[0] == '/')
            {
                return new string(target[1..]);
            }
            if (target.StartsWith("xl/"))
            {
                return new string(target);
            }
            return string.Concat("xl/", target);
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
                ParseShared(ReadAll(entry));
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
                ParseShared(await ReadAllAsync(entry, ct).ConfigureAwait(false));
            }
        }

        private void ParseShared(ReadOnlySpan<byte> src)
        {
            // Decoded text is never longer than its XML, so src.Length bounds the flat buffer.
            _sharedFlat = ArrayPool<byte>.Shared.Rent(Math.Max(1, src.Length));

            // Pre-size offsets from <sst uniqueCount="N"> to avoid repeated List growth.
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

            var offsets = new List<int>(uniqueCount + 1) { 0 };
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
                offsets.Add(flat);
            }
            _sharedOffsets = [.. offsets];
            _sharedCount = offsets.Count - 1;
        }

        // Whole-part bytes, or null when the part is absent. Sync and async variants share every parser.
        private static byte[]? Bytes(ZipArchive zip, string name)
        {
            var entry = zip.GetEntry(name);
            return entry is null ? null : ReadAll(entry);
        }

        private static async ValueTask<byte[]?> BytesAsync(ZipArchive zip, string name, CancellationToken ct)
        {
            var entry = zip.GetEntry(name);
            return entry is null ? null : await ReadAllAsync(entry, ct).ConfigureAwait(false);
        }

        private static byte[] ReadAll(ZipArchiveEntry entry)
        {
            var buf = new byte[entry.Length];
            using var s = entry.Open();
            s.ReadExactly(buf);
            return buf;
        }

        private static async ValueTask<byte[]> ReadAllAsync(ZipArchiveEntry entry, CancellationToken ct)
        {
            var buf = new byte[entry.Length];
            var s = await entry.OpenAsync(ct).ConfigureAwait(false);
            await using (s.ConfigureAwait(false))
            {
                await s.ReadExactlyAsync(buf.AsMemory(), ct).ConfigureAwait(false);
            }
            return buf;
        }

        private static string Decode(ReadOnlySpan<byte> src)
        {
            if (src.IsEmpty)
            {
                return string.Empty;
            }
            Span<byte> dest = src.Length <= 256 ? stackalloc byte[src.Length] : new byte[src.Length];
            int w = XlsxXml.Decode(src, dest);
            return System.Text.Encoding.UTF8.GetString(dest[..w]);
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

        // Yields each open-tag span (including '<' and '>') whose start matches `prefix`, over a full buffer.
        // Returns ReadOnlyMemory<byte> slices — no per-tag allocation.
        private static IEnumerable<ReadOnlyMemory<byte>> Tags(ReadOnlyMemory<byte> buf, byte[] prefix)
        {
            int pos = 0;
            while (true)
            {
                int start = IdxOf(buf.Span, pos, prefix);
                if (start < 0)
                {
                    yield break;
                }
                int end = IdxOf(buf.Span, start, (byte)'>');
                if (end < 0)
                {
                    yield break;
                }
                yield return buf.Slice(start, end - start + 1);
                pos = end + 1;
            }
        }
    }
}
