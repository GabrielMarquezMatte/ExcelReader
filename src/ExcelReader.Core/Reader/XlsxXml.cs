using System.Buffers;
using System.Buffers.Text;
using System.Runtime.CompilerServices;
using System.Text;

namespace ExcelReader.Core.Reader
{
    internal static class XlsxXml
    {
        internal static int ParseIntOr(ReadOnlySpan<byte> source, int fallback)
        {
            return Utf8Parser.TryParse(source, out int value, out _) ? value : fallback;
        }

        public static ReadOnlySpan<byte> Attr(ReadOnlySpan<byte> openTag, ReadOnlySpan<byte> name)
        {
            int i = openTag.IndexOf(name);
            if (i < 0)
            {
                return default;
            }
            int start = i + name.Length;
            if (start >= openTag.Length)
            {
                return default;
            }
            byte quote = openTag[start];
            if (quote is not ((byte)'"' or (byte)'\''))
            {
                return default;
            }
            start++;
            int end = openTag[start..].IndexOf(quote);
            return end < 0 ? default : openTag.Slice(start, end);
        }

        public static int ColumnIndex(ReadOnlySpan<byte> cellRef)
        {
            if (cellRef.IsEmpty)
            {
                return -1;
            }
            uint a = (uint)(cellRef[0] - 'A');
            if (a > 25)
            {
                return -1;
            }
            if (cellRef.Length == 1)
            {
                return (int)a;
            }
            uint b = (uint)(cellRef[1] - 'A');
            if (b > 25)
            {
                return (int)a;
            }
            if (cellRef.Length == 2)
            {
                return (int)(((a + 1) * 26) + b);
            }
            uint c = (uint)(cellRef[2] - 'A');
            if (c > 25)
            {
                return (int)(((a + 1) * 26) + b);
            }
            if (cellRef.Length == 3 || (uint)(cellRef[3] - 'A') > 25)
            {
                return (int)(((((a + 1) * 26) + b + 1) * 26) + c);
            }
            return ExcelLimits.MaxColumns;
        }

        public static int Decode(ReadOnlySpan<byte> src, Span<byte> dest)
        {
            int w = 0;
            while (!src.IsEmpty)
            {
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
                    src[..(semi + 1)].CopyTo(dest[w..]);
                    w += semi + 1;
                }
                src = src[(semi + 1)..];
            }
            return w;
        }

        private static int DecodeNumeric(ReadOnlySpan<byte> body, Span<byte> dest, ReadOnlySpan<byte> raw)
        {
            const int MaxCodepoint = 0x10FFFF;
            int cp = 0;
            bool ok = false;
            if (body.Length > 0 && (body[0] == 'x' || body[0] == 'X'))
            {
                foreach (ref readonly byte d in body[1..])
                {
                    int v = HexVal(d);
                    if (v < 0 || cp > MaxCodepoint) { ok = false; break; }
                    cp = (cp * 16) + v;
                    ok = true;
                }
            }
            else
            {
                foreach (ref readonly byte d in body)
                {
                    if (d is < (byte)'0' or > (byte)'9' || cp > MaxCodepoint) { ok = false; break; }
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

        internal static int WriteTextRuns(ReadOnlySpan<byte> si, Span<byte> dest)
        {
            return WriteTextRuns(si, dest, "<t"u8, "</t>"u8, "<rPh"u8, "</rPh>"u8);
        }

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
                    remaining = remaining[(rPhEnd + rPhClose.Length)..];
                    continue;
                }
                remaining = remaining[(tIndex + tOpen.Length)..];
                var openIndex = remaining.IndexOf((byte)'>');
                if (openIndex < 0)
                {
                    break;
                }
                if (openIndex > 0 && remaining[openIndex - 1] == '/')
                {
                    remaining = remaining[(openIndex + 1)..];
                    continue;
                }
                remaining = remaining[(openIndex + 1)..];
                var closeIndex = remaining.IndexOf(tClose);
                if (closeIndex < 0)
                {
                    break;
                }
                var innerText = remaining[..closeIndex];
                var written = Decode(innerText, destSlice);
                totalWritten += written;
                destSlice = destSlice[written..];
                remaining = remaining[(closeIndex + tClose.Length)..];
            }
            return totalWritten;
        }

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
                        return default;
                    }
                    if (b == (byte)':')
                    {
                        return src.Slice(nameStart, j - nameStart + 1);
                    }
                }
                return default;
            }
        }

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

        [SkipLocalsInit]
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

        internal static Dictionary<string, string> ParseRelationships(ReadOnlySpan<byte> relsBytes)
        {
            Dictionary<string, string> rels = new(StringComparer.Ordinal);
            if (relsBytes.IsEmpty)
            {
                return rels;
            }
            foreach (ReadOnlySpan<byte> tag in new TagSpanEnumerable(relsBytes, "<Relationship"u8))
            {
                string id = DecodeToString(Attr(tag, " Id="u8));
                if (id.Length > 0)
                {
                    rels[id] = DecodeToString(Attr(tag, " Target="u8));
                }
            }
            return rels;
        }
    }
}
