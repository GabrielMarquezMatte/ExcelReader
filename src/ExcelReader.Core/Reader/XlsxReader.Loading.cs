using System.Buffers;
using System.IO.Compression;

namespace ExcelReader.Core.Reader
{
    public sealed partial class XlsxReader
    {
        // --- workbook / shared-strings loading (one-time, small except sharedStrings) ---

        private static (string Name, string Path)[] ParseSheets(byte[]? wbBytes, byte[]? relsBytes)
        {
            if (wbBytes is null)
            {
                return [];
            }
            // rId -> target part path
            Dictionary<string, string> rels = new(StringComparer.Ordinal);
            if (relsBytes is not null)
            {
                foreach (var tag in Tags(relsBytes, "<Relationship"u8.ToArray()))
                {
                    var id = Decode(XlsxXml.Attr(tag, " Id=\""u8));
                    var target = Decode(XlsxXml.Attr(tag, " Target=\""u8));
                    if (id.Length > 0)
                    {
                        rels[id] = target;
                    }
                }
            }
            var sheets = new List<(string, string)>();
            foreach (var tag in Tags(wbBytes, "<sheet "u8.ToArray()))
            {
                var name = Decode(XlsxXml.Attr(tag, " name=\""u8));
                var rid = Decode(XlsxXml.Attr(tag, " r:id=\""u8));
                if (rels.TryGetValue(rid, out var target))
                {
                    sheets.Add((name, NormalizePart(target)));
                }
            }
            return [.. sheets];
        }

        private static string NormalizePart(string target)
        {
            if (target.StartsWith('/'))
            {
                return target[1..];
            }
            if (target.StartsWith("xl/", StringComparison.Ordinal))
            {
                return target;
            }
            return "xl/" + target;
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

        private void ParseShared(byte[] src)
        {
            // Decoded text is never longer than its XML, so src.Length bounds the flat buffer.
            _sharedFlat = ArrayPool<byte>.Shared.Rent(Math.Max(1, src.Length));
            var offsets = new List<int> { 0 };
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
                    flat += XlsxXml.WriteTextRuns(src.AsSpan(open + 1, end - open - 1), _sharedFlat.AsSpan(flat));
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
        private static IEnumerable<byte[]> Tags(byte[] buf, byte[] prefix)
        {
            int pos = 0;
            while (true)
            {
                int start = IdxOf(buf, pos, prefix);
                if (start < 0)
                {
                    yield break;
                }
                int end = IdxOf(buf, start, (byte)'>');
                if (end < 0)
                {
                    yield break;
                }
                yield return buf[start..(end + 1)];
                pos = end + 1;
            }
        }
    }
}
