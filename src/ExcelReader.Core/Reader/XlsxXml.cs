using System.Buffers;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelReader.Core.Reader
{
    // Byte-scan + XML-entity-decode primitives for the tiny SpreadsheetML subset we read.
    // Everything works on UTF-8 ReadOnlySpan<byte> so cell values never round-trip through string.
    internal static class XlsxXml
    {
        // Returns the value of attribute `name` inside an open tag span like `<c r="A1" s="1" t="s">`.
        // `name` must include the leading space and `="`, e.g. " t=\"" — the leading space gives a cheap
        // word boundary so " r=\"" doesn't match inside another attribute. Empty span if absent.
        public static ReadOnlySpan<byte> Attr(ReadOnlySpan<byte> openTag, ReadOnlySpan<byte> name)
        {
            int i = openTag.IndexOf(name);
            if (i < 0)
            {
                return default;
            }
            int start = i + name.Length;
            int end = openTag[start..].IndexOf((byte)'"');
            return end < 0 ? default : openTag.Slice(start, end);
        }

        // Column reference letters (the "B" in "B2") -> 0-based column index. "" -> -1.
        public static int ColumnIndex(ReadOnlySpan<byte> cellRef)
        {
            ref var c0 = ref MemoryMarshal.GetReference(cellRef);
            int col = 0;
            int i = 0;
            var length = cellRef.Length;
            for (; i < length; i++)
            {
                var c = Unsafe.Add(ref c0, i);
                uint val = (uint)(c - 'A');
                if (val > 25)
                {
                    break;
                }
                col = (col * 26) + (int)val + 1;
            }
            return i == 0 ? -1 : col - 1;
        }

        // Decode the 5 predefined XML entities + numeric &#d;/&#xH; from `src` into `dest`,
        // returning bytes written. Decoded output is never longer than `src`, so a dest sized to
        // src.Length always fits. Unknown entities are copied through literally.
        public static int Decode(ReadOnlySpan<byte> src, Span<byte> dest)
        {
            int w = 0;
            while (!src.IsEmpty)
            {
                // The literal run up to the next '&' or '<' is bulk-copied; IndexOfAny and CopyTo
                // are both SIMD-vectorized, so plain text still costs one scan + one copy.
                int special = src.IndexOfAny((byte)'&', (byte)'<');
                if (special < 0)
                {
                    src.CopyTo(dest[w..]);
                    return w + src.Length;
                }
                src[..special].CopyTo(dest[w..]);
                w += special;
                src = src[special..];

                if (src[0] == (byte)'<')
                {
                    if (!src.StartsWith("<![CDATA["u8))
                    {
                        dest[w++] = src[0];
                        src = src[1..];
                        continue;
                    }

                    src = src[9..];
                    int end = src.IndexOf("]]>"u8);
                    if (end < 0)
                    {
                        src.CopyTo(dest[w..]);
                        return w + src.Length;
                    }
                    src[..end].CopyTo(dest[w..]);
                    w += end;
                    src = src[(end + 3)..];
                    continue;
                }

                int semi = src.IndexOf((byte)';');
                if (semi < 0)
                {
                    // No terminator anywhere in what remains, so no entity can follow — copy it all.
                    src.CopyTo(dest[w..]);
                    return w + src.Length;
                }
                var ent = src[1..semi];
                if (ent.SequenceEqual("amp"u8)) { dest[w++] = (byte)'&'; }
                else if (ent.SequenceEqual("lt"u8)) { dest[w++] = (byte)'<'; }
                else if (ent.SequenceEqual("gt"u8)) { dest[w++] = (byte)'>'; }
                else if (ent.SequenceEqual("quot"u8)) { dest[w++] = (byte)'"'; }
                else if (ent.SequenceEqual("apos"u8)) { dest[w++] = (byte)'\''; }
                else if (ent.Length > 1 && ent[0] == '#')
                {
                    w += DecodeNumeric(ent[1..], dest[w..], src[..(semi + 1)]);
                }
                else
                {
                    // Unknown entity: copy the raw "&...;" through unchanged.
                    src[..(semi + 1)].CopyTo(dest[w..]);
                    w += semi + 1;
                }
                src = src[(semi + 1)..];
            }
            return w;
        }

        // `body` is the part after '#': decimal digits, or 'x'/'X' + hex. `raw` is the whole "&#..;"
        // used as the literal fallback when the codepoint is malformed.
        private static int DecodeNumeric(ReadOnlySpan<byte> body, Span<byte> dest, ReadOnlySpan<byte> raw)
        {
            int cp = 0;
            bool ok = false;
            if (body.Length > 0 && (body[0] == 'x' || body[0] == 'X'))
            {
                foreach (ref readonly byte d in body[1..])
                {
                    int v = HexVal(d);
                    if (v < 0) { ok = false; break; }
                    cp = (cp * 16) + v;
                    ok = true;
                }
            }
            else
            {
                foreach (ref readonly byte d in body)
                {
                    if (d is < (byte)'0' or > (byte)'9') { ok = false; break; }
                    cp = (cp * 10) + (d - '0');
                    ok = true;
                }
            }
            if (ok && Rune.IsValid(cp) && new Rune(cp).TryEncodeToUtf8(dest, out int written))
            {
                return written;
            }
            raw.CopyTo(dest);
            return raw.Length;
        }

        // Scans every <t>...</t> run inside `si`, entity-decodes each one, and writes the result
        // into `dest` starting at offset 0. Returns total bytes written. Text inside <rPh> (phonetic
        // guide runs, e.g. Japanese furigana) is skipped — it's a pronunciation hint, not cell content.
        // `dest` must be at least `si.Length` bytes — decoded text is never longer than its source XML.
        internal static int WriteTextRuns(ReadOnlySpan<byte> si, Span<byte> dest)
        {
            return WriteTextRuns(si, dest, "<t"u8, "</t>"u8, "<rPh"u8, "</rPh>"u8);
        }

        // Namespace-prefixed overload: the element tokens (e.g. "<x:t", "</x:t>") are supplied by the
        // caller, which built them once from the part's root-element prefix (see NsTokens / DetectElementPrefix).
        internal static int WriteTextRuns(ReadOnlySpan<byte> si, Span<byte> dest,
            ReadOnlySpan<byte> tOpen, ReadOnlySpan<byte> tClose, ReadOnlySpan<byte> rPhOpen, ReadOnlySpan<byte> rPhClose)
        {
            int totalWritten = 0;
            ReadOnlySpan<byte> remaining = si;
            Span<byte> destSlice = dest;
            while (true)
            {
                var tIndex = remaining.IndexOf(tOpen);
                if (tIndex < 0)
                {
                    break;
                }
                var rPhIndex = remaining.IndexOf(rPhOpen);
                if (rPhIndex >= 0 && rPhIndex < tIndex)
                {
                    var rPhEnd = remaining.IndexOf(rPhClose);
                    if (rPhEnd < 0)
                    {
                        break;
                    }
                    remaining = remaining[(rPhEnd + rPhClose.Length)..]; // Skip past "</rPh>"
                    continue;
                }
                remaining = remaining[(tIndex + tOpen.Length)..]; // Skip past "<t"
                var openIndex = remaining.IndexOf((byte)'>');
                if (openIndex < 0)
                {
                    break;
                }
                if (openIndex > 0 && remaining[openIndex - 1] == '/')
                {
                    remaining = remaining[(openIndex + 1)..]; // Skip past the self-closing tag
                    continue;
                }
                remaining = remaining[(openIndex + 1)..]; // Skip past the opening tag
                var closeIndex = remaining.IndexOf(tClose);
                if (closeIndex < 0)
                {
                    break;
                }
                var innerText = remaining[..closeIndex];
                var written = Decode(innerText, destSlice);
                totalWritten += written;
                destSlice = destSlice[written..];
                remaining = remaining[(closeIndex + tClose.Length)..]; // Skip past the closing tag
            }
            return totalWritten;
        }

        // The element-name prefix on the document's root element (e.g. "x:" in <x:worksheet ...>),
        // including the trailing ':'. Empty span when elements are unprefixed — the default-namespace
        // case Excel and most producers emit. Skips the XML declaration, comments, and DOCTYPE.
        internal static ReadOnlySpan<byte> DetectElementPrefix(ReadOnlySpan<byte> src)
        {
            int i = 0;
            while (true)
            {
                int lt = src[i..].IndexOf((byte)'<');
                if (lt < 0)
                {
                    return default;
                }
                i += lt + 1;
                if (i >= src.Length)
                {
                    return default;
                }
                // '?' = <?xml?>, '!' = <!-- --> / <!DOCTYPE>, '/' = stray close tag — none are the root.
                if (src[i] is (byte)'?' or (byte)'!' or (byte)'/')
                {
                    continue;
                }
                int nameStart = i;
                for (int j = i; j < src.Length; j++)
                {
                    byte b = src[j];
                    if (b is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n' or (byte)'>' or (byte)'/')
                    {
                        return default; // root element name has no ':' — unprefixed
                    }
                    if (b == (byte)':')
                    {
                        return src.Slice(nameStart, j - nameStart + 1); // "x:" (colon included)
                    }
                }
                return default;
            }
        }

        // Builds a literal element token "lead + prefix + rest", e.g. ("</", "x:", "row") -> "</x:row".
        internal static byte[] Token(ReadOnlySpan<byte> lead, ReadOnlySpan<byte> prefix, ReadOnlySpan<byte> rest)
        {
            byte[] token = new byte[lead.Length + prefix.Length + rest.Length];
            lead.CopyTo(token);
            prefix.CopyTo(token.AsSpan(lead.Length));
            rest.CopyTo(token.AsSpan(lead.Length + prefix.Length));
            return token;
        }
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static int HexVal(byte d)
        {
            var value = (d & 0xF) + (9 * (d >> 6));
            var isNum = (d - '0') <= 9u;
            var lower = (d | 0x20) - 'a';
            var isAlpha = lower <= 5u;
            return (isNum || isAlpha) ? value : -1;
        }

        // Decode an XML-entity-encoded attribute/text value straight to a string. Small values (the
        // common case: rIds, sheet names, part paths) use a stack buffer; larger ones use a pooled array.
        internal static string DecodeToString(ReadOnlySpan<byte> src)
        {
            if (src.IsEmpty)
            {
                return string.Empty;
            }
            if (src.Length <= 256)
            {
                Span<byte> dest = stackalloc byte[src.Length];
                int w = Decode(src, dest);
                return Encoding.UTF8.GetString(dest[..w]);
            }
            byte[] buffer = ArrayPool<byte>.Shared.Rent(src.Length);
            try
            {
                int w = Decode(src, buffer);
                return Encoding.UTF8.GetString(buffer.AsSpan(0, w));
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }

        // "/xl/worksheets/sheet1.xml" -> "xl/worksheets/sheet1.xml"; a bare "worksheets/sheet1.xml"
        // gets the "xl/" prefix that OPC part paths in workbook rels are always relative to.
        internal static string NormalizePart(ReadOnlySpan<char> target)
        {
            if (target.Length > 0 && target[0] == '/')
            {
                return new string(target[1..]);
            }
            if (target.StartsWith("xl/"))
            {
                return new string(target);
            }
            return $"xl/{target}";
        }

        // Parses a `.rels` part (a list of <Relationship Id="rIdN" Target="..."/> elements) into an
        // rId -> target-part-path lookup. Shared by the XLSX and XLSB loaders — both use the same OPC
        // relationships schema.
        internal static Dictionary<string, string> ParseRelationships(ReadOnlySpan<byte> relsBytes)
        {
            Dictionary<string, string> rels = new(StringComparer.Ordinal);
            if (relsBytes.IsEmpty)
            {
                return rels;
            }
            foreach (ReadOnlySpan<byte> tag in new TagSpanEnumerable(relsBytes, "<Relationship"u8))
            {
                string id = DecodeToString(Attr(tag, " Id=\""u8));
                if (id.Length > 0)
                {
                    rels[id] = DecodeToString(Attr(tag, " Target=\""u8));
                }
            }
            return rels;
        }
    }
}
