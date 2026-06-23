using System.Buffers;
using System.Buffers.Text;
using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using ExcelReader.Core.Enums;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
{
    public sealed class XlsxReader : IDisposable, IAsyncDisposable
    {
        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private readonly ZipArchive _zip;
        private readonly (string Name, string Path)[] _sheets;
        private readonly BitArray _styleIsDate; // cellXfs index -> true when that style renders as a date/time
        private int _current;

        private byte[] _sharedFlat = [];      // pooled; all decoded shared-string bytes concatenated
        private int[] _sharedOffsets = [0];   // string i = _sharedFlat[_offsets[i].._offsets[i+1]]
        private int _sharedCount;
        private bool _sharedLoaded;

        // Sync open: reads the central directory and workbook/styles parts synchronously.
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP003:Dispose previous before re-assigning",
            Justification = "Readonly field, first and only assignment in this constructor.")]
        internal XlsxReader(Stream stream, bool leaveOpen)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            try
            {
                _sheets = ParseSheets(Bytes(_zip, "xl/workbook.xml"), Bytes(_zip, "xl/_rels/workbook.xml.rels"));
                if (_sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                _styleIsDate = ParseStyleDateFlags(Bytes(_zip, "xl/styles.xml"));
            }
            catch
            {
                _zip.Dispose();
                if (!leaveOpen)
                {
                    stream.Dispose();
                }
                throw;
            }
        }

        private XlsxReader(Stream stream, bool leaveOpen, ZipArchive zip,
            (string Name, string Path)[] sheets, BitArray styleIsDate)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _zip = zip;
            _sheets = sheets;
            _styleIsDate = styleIsDate;
        }

        // Async open: central directory and parts are read with the .NET 10 async zip APIs.
        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP001:Dispose created",
            Justification = "zip ownership transfers to the returned reader; disposed there or in the catch.")]
        internal static async ValueTask<XlsxReader> CreateAsync(Stream stream, bool leaveOpen, CancellationToken ct)
        {
            ZipArchive? zip = null;
            try
            {
                zip = await ZipArchive.CreateAsync(stream, ZipArchiveMode.Read, leaveOpen: true, entryNameEncoding: null, ct).ConfigureAwait(false);
                var wb = await BytesAsync(zip, "xl/workbook.xml", ct).ConfigureAwait(false);
                var rels = await BytesAsync(zip, "xl/_rels/workbook.xml.rels", ct).ConfigureAwait(false);
                var sheets = ParseSheets(wb, rels);
                if (sheets.Length == 0)
                {
                    throw new InvalidDataException("The workbook contains no sheets.");
                }
                var styleIsDate = ParseStyleDateFlags(await BytesAsync(zip, "xl/styles.xml", ct).ConfigureAwait(false));
                return new XlsxReader(stream, leaveOpen, zip, sheets, styleIsDate);
            }
            catch
            {
                if (zip is not null)
                {
                    await zip.DisposeAsync().ConfigureAwait(false);
                }
                if (!leaveOpen)
                {
                    await stream.DisposeAsync().ConfigureAwait(false);
                }
                throw;
            }
        }

        public string SheetName => _sheets[_current].Name;
        public int SheetCount => _sheets.Length;

        // A numeric cell whose style index maps to a date/time format is reported as CellType.Date.
        internal bool IsDateStyle(int style)
        {
            return (uint)style < (uint)_styleIsDate.Length && _styleIsDate[style];
        }


        public bool TryMoveToSheet(string name)
        {
            for (int i = 0; i < _sheets.Length; i++)
            {
                if (string.Equals(_sheets[i].Name, name, StringComparison.Ordinal))
                {
                    _current = i;
                    return true;
                }
            }
            return false;
        }

        public void MoveToSheet(int index)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, _sheets.Length);
            _current = index;
        }

        [SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "Enumerator is a class so the same type can also expose MoveNextAsync for the async path.")]
        public Enumerator GetEnumerator()
        {
            EnsureSharedLoaded();
            var entry = _zip.GetEntry(_sheets[_current].Path)
                ?? throw new InvalidDataException($"Worksheet part not found: {_sheets[_current].Path}");
            return new Enumerator(this, entry.Open());
        }

        /// <summary>
        /// Streaming async enumerator over the current sheet. Use with a manual loop — <c>Current</c>
        /// is a ref struct (<c>Row</c>), so <c>await foreach</c> cannot bind it:
        /// <code>
        /// await using var e = await reader.GetAsyncEnumeratorAsync(ct);
        /// while (await e.MoveNextAsync()) { var row = e.Current; /* ... */ }
        /// </code>
        /// </summary>
        public async ValueTask<Enumerator> GetAsyncEnumeratorAsync(CancellationToken ct = default)
        {
            await EnsureSharedLoadedAsync(ct).ConfigureAwait(false);
            var entry = _zip.GetEntry(_sheets[_current].Path)
                ?? throw new InvalidDataException($"Worksheet part not found: {_sheets[_current].Path}");
            var sheet = await entry.OpenAsync(ct).ConfigureAwait(false);
            return new Enumerator(this, sheet, ct);
        }

        internal ReadOnlySpan<byte> SharedSpan => _sharedFlat;

        internal (int Start, int Length) SharedAt(int index)
        {
            if ((uint)index >= (uint)_sharedCount)
            {
                return (0, 0);
            }
            return (_sharedOffsets[index], _sharedOffsets[index + 1] - _sharedOffsets[index]);
        }

        public void Dispose()
        {
            if (_sharedFlat.Length > 0)
            {
                ArrayPool<byte>.Shared.Return(_sharedFlat);
                _sharedFlat = [];
            }
            _zip.Dispose();
            if (!_leaveOpen)
            {
                _stream.Dispose();
            }
        }

        public async ValueTask DisposeAsync()
        {
            if (_sharedFlat.Length > 0)
            {
                ArrayPool<byte>.Shared.Return(_sharedFlat);
                _sharedFlat = [];
            }
            await _zip.DisposeAsync().ConfigureAwait(false);
            if (!_leaveOpen)
            {
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
        }

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

        // Builds the cellXfs-index -> isDate table from xl/styles.xml. A style is a date when its
        // numFmtId is a builtin date/time format or a custom <numFmt> whose code reads as a date.
        // ponytail: 1900 date system only (see Cell.TryGetDateTime). Add date1904 handling if needed.
        private static BitArray ParseStyleDateFlags(byte[]? src)
        {
            if (src is null)
            {
                return new BitArray(0);
            }

            // Custom formats: numFmtId -> isDate(formatCode). Builtin ids (<164) handled by IsBuiltinDate.
            Dictionary<int, bool> custom = new(capacity: 16);
            foreach (var tag in Tags(src, "<numFmt "u8.ToArray()))
            {
                int id = ParseIntOr(XlsxXml.Attr(tag, " numFmtId=\""u8), -1);
                if (id >= 0)
                {
                    custom[id] = LooksLikeDateFormat(Decode(XlsxXml.Attr(tag, " formatCode=\""u8)));
                }
            }

            // Only the <xf> entries inside <cellXfs> are cell styles; <cellStyleXfs> is the master table.
            int region = IdxOf(src, 0, "<cellXfs"u8);
            if (region < 0)
            {
                return new BitArray(0);
            }
            int open = IdxOf(src, region, (byte)'>');
            int end = IdxOf(src, open, "</cellXfs>"u8);
            if (open < 0 || end < 0)
            {
                return new BitArray(0);
            }
            List<bool> flags = new(capacity: 16);
            foreach (var xf in Tags(src[(open + 1)..end], "<xf "u8.ToArray()))
            {
                int numFmtId = ParseIntOr(XlsxXml.Attr(xf, " numFmtId=\""u8), 0);
                flags.Add(custom.TryGetValue(numFmtId, out bool d) ? d : IsBuiltinDateFormat(numFmtId));
            }
            return new BitArray(flags.ToArray());
        }

        // Builtin SpreadsheetML date/time numFmtIds (ECMA-376 §18.8.30, incl. locale variants).
        private static bool IsBuiltinDateFormat(int id)
        {
            return id is (>= 14 and <= 22) or (>= 27 and <= 36) or (>= 45 and <= 47) or (>= 50 and <= 58) or (>= 71 and <= 81);
        }

        // True if a format code contains a date/time token (y/m/d/h/s) outside quoted text, [bracketed]
        // sections, and \-escapes. ponytail: heuristic, not a full format parser — upgrade if a format
        // with date letters only inside literals is misclassified.

        private static bool LooksLikeDateFormat(string code)
        {
            int i = 0;
            while (i < code.Length)
            {
                switch (code[i])
                {
                    case '"':
                        i = code.IndexOf('"', i + 1);
                        if (i < 0) { return false; }
                        i++;
                        break;
                    case '[':
                        i = code.IndexOf(']', i + 1);
                        if (i < 0) { return false; }
                        i++;
                        break;
                    case '\\':
                        i += 2; // skip the escaped char
                        break;
                    case 'y' or 'Y' or 'm' or 'M' or 'd' or 'D' or 'h' or 'H' or 's' or 'S':
                        return true;
                    default:
                        i++;
                        break;
                }
            }
            return false;
        }

        private static int ParseIntOr(ReadOnlySpan<byte> src, int fallback)
        {
            return Utf8Parser.TryParse(src, out int v, out _) ? v : fallback;
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

        // Forward-only, low-memory worksheet scanner. Streams the sheet through a refillable pooled
        // buffer; a single <c>...</c> element is guaranteed contiguous (the buffer grows if needed).
        [SuppressMessage("Design", "CA1034:Nested types should not be visible",
            Justification = "Public nested enumerator is the standard foreach pattern.")]
        public sealed class Enumerator : IDisposable, IAsyncDisposable
        {
            private const int InitialBuf = 64 * 1024;
            private const int InitialVals = 4 * 1024;
            private const int InitialCells = 32;

            // Borrowed: the reader outlives the enumerator and owns its own disposal — do not dispose here.
            [SuppressMessage("Usage", "CA2213:Disposable fields should be disposed", Justification = "Borrowed, not owned.")]
            [SuppressMessage("SharpSource", "SS066:Disposable field is not disposed", Justification = "Borrowed, not owned.")]
            private readonly XlsxReader _reader;
            private readonly CancellationToken _ct; // honored only by the async path
            // Owned: opened by Get(Async)Enumerator for this enumerator alone; disposed in Dispose(Async).
            [SuppressMessage("SharpSource", "SS066:Disposable field is not disposed", Justification = "Disposed in Dispose().")]
            private Stream? _sheet;
            private byte[] _buf;
            private int _pos;
            private int _len;
            private bool _eof;

            private byte[] _vals;     // decoded values for the current row (numbers/inline/formula text)
            private int _valLen;
            private CellDesc[] _cells; // current row's cells, ascending by column
            private int _cellCount;
            private int _nextCol;

            internal Enumerator(XlsxReader reader, Stream sheet, CancellationToken ct = default)
            {
                _reader = reader;
                _sheet = sheet;
                _ct = ct;
                _buf = ArrayPool<byte>.Shared.Rent(InitialBuf);
                _vals = ArrayPool<byte>.Shared.Rent(InitialVals);
                _cells = ArrayPool<CellDesc>.Shared.Rent(InitialCells);
            }

            public Row Current =>
                new(_cells.AsSpan(0, _cellCount), _vals.AsSpan(0, _valLen), _reader.SharedSpan);

            // The sync and async row scanners share every span-touching helper below (ClassifyHead,
            // ClassifyRowHead, ReadCellOpenTag, EmitCell). The two families differ only in whether the
            // buffer refill (Fill / FillAsync) and the byte searches await — so the span work is factored
            // into sync helpers that never hold a span across an await.

            public bool MoveNext()
            {
                while (true)
                {
                    int lt = IndexOf((byte)'<');
                    if (lt < 0)
                    {
                        return false;
                    }
                    _pos = lt;
                    Ensure(12);
                    switch (ClassifyHead())
                    {
                        case HeadKind.End:
                            return false;
                        case HeadKind.Row:
                            BeginRow(out bool selfClose);
                            if (!selfClose)
                            {
                                ParseRow();
                            }
                            return true;
                        default:
                            int skip = IndexOf((byte)'>');
                            if (skip < 0)
                            {
                                return false;
                            }
                            _pos = skip + 1;
                            break;
                    }
                }
            }

            public async ValueTask<bool> MoveNextAsync()
            {
                while (true)
                {
                    int lt = await IndexOfAsync((byte)'<').ConfigureAwait(false);
                    if (lt < 0)
                    {
                        return false;
                    }
                    _pos = lt;
                    await EnsureAsync(12).ConfigureAwait(false);
                    switch (ClassifyHead())
                    {
                        case HeadKind.End:
                            return false;
                        case HeadKind.Row:
                            BeginRow(out bool selfClose);
                            if (!selfClose)
                            {
                                await ParseRowAsync().ConfigureAwait(false);
                            }
                            return true;
                        default:
                            int skip = await IndexOfAsync((byte)'>').ConfigureAwait(false);
                            if (skip < 0)
                            {
                                return false;
                            }
                            _pos = skip + 1;
                            break;
                    }
                }
            }

            // Consumes the <row ...> open tag and resets per-row state. Call only after ClassifyHead()==Row.
            private void BeginRow(out bool selfClose)
            {
                int gt = IndexOf((byte)'>'); // open tag already fully buffered by the Ensure(12) above
                selfClose = _buf[gt - 1] == '/';
                _pos = gt + 1;
                _cellCount = 0;
                _valLen = 0;
                _nextCol = 0;
            }

            private void ParseRow()
            {
                while (true)
                {
                    int lt = IndexOf((byte)'<');
                    if (lt < 0)
                    {
                        return;
                    }
                    _pos = lt;
                    Ensure(6);
                    switch (ClassifyRowHead())
                    {
                        case RowHead.EndRow:
                            int gt = IndexOf((byte)'>');
                            _pos = gt < 0 ? _len : gt + 1;
                            return;
                        case RowHead.Cell:
                            ParseCell();
                            break;
                        default:
                            int skip = IndexOf((byte)'>');
                            if (skip < 0)
                            {
                                return;
                            }
                            _pos = skip + 1;
                            break;
                    }
                }
            }

            private async ValueTask ParseRowAsync()
            {
                while (true)
                {
                    int lt = await IndexOfAsync((byte)'<').ConfigureAwait(false);
                    if (lt < 0)
                    {
                        return;
                    }
                    _pos = lt;
                    await EnsureAsync(6).ConfigureAwait(false);
                    switch (ClassifyRowHead())
                    {
                        case RowHead.EndRow:
                            int gt = await IndexOfAsync((byte)'>').ConfigureAwait(false);
                            _pos = gt < 0 ? _len : gt + 1;
                            return;
                        case RowHead.Cell:
                            await ParseCellAsync().ConfigureAwait(false);
                            break;
                        default:
                            int skip = await IndexOfAsync((byte)'>').ConfigureAwait(false);
                            if (skip < 0)
                            {
                                return;
                            }
                            _pos = skip + 1;
                            break;
                    }
                }
            }

            private void ParseCell()
            {
                int gt = IndexOf((byte)'>'); // end of the <c ...> open tag (buffered, no shift yet)
                var header = ReadCellOpenTag(gt);
                if (header.SelfClose)
                {
                    return; // empty cell — store nothing; _pos already past the tag
                }
                int cEnd = IndexOfSeq("</c>"u8); // ensures whole cell contiguous; shifts _pos consistently
                if (cEnd < 0)
                {
                    _pos = _len;
                    return;
                }
                EmitCell(header.Kind, _buf.AsSpan(_pos, cEnd - _pos), header.Col, header.Style);
                _pos = cEnd + 4;
            }

            private async ValueTask ParseCellAsync()
            {
                int gt = await IndexOfAsync((byte)'>').ConfigureAwait(false);
                var header = ReadCellOpenTag(gt); // sync: no span survives into the next await
                if (header.SelfClose)
                {
                    return;
                }
                int cEnd = await IndexOfCloseCellAsync().ConfigureAwait(false);
                if (cEnd < 0)
                {
                    _pos = _len;
                    return;
                }
                EmitCell(header.Kind, _buf.AsSpan(_pos, cEnd - _pos), header.Col, header.Style);
                _pos = cEnd + 4;
            }

            private enum HeadKind { End, Row, Skip }

            private HeadKind ClassifyHead()
            {
                var head = _buf.AsSpan(_pos, Math.Min(12, _len - _pos));
                if (head.StartsWith("</sheetData"u8) || head.StartsWith("</worksheet"u8))
                {
                    return HeadKind.End;
                }
                if (head.StartsWith("<row"u8) && (head.Length < 5 || IsBoundary(head[4])))
                {
                    return HeadKind.Row;
                }
                return HeadKind.Skip;
            }

            private enum RowHead { EndRow, Cell, Other }

            private RowHead ClassifyRowHead()
            {
                var head = _buf.AsSpan(_pos, Math.Min(6, _len - _pos));
                if (head.StartsWith("</row"u8))
                {
                    return RowHead.EndRow;
                }
                if (head.StartsWith("<c"u8) && (head.Length < 3 || IsBoundary(head[2])))
                {
                    return RowHead.Cell;
                }
                return RowHead.Other;
            }

            private readonly record struct CellHeader(int Col, int Style, Kind Kind, bool SelfClose);

            // Parses the <c ...> open tag ending at `gt`, advances _pos past it, and returns the extracted
            // (non-span) attributes. Holds spans only within this synchronous call, so callers may await after.
            private CellHeader ReadCellOpenTag(int gt)
            {
                var open = _buf.AsSpan(_pos, gt - _pos + 1);
                var rRef = XlsxXml.Attr(open, " r=\""u8);
                var sVal = XlsxXml.Attr(open, " s=\""u8);
                var tVal = XlsxXml.Attr(open, " t=\""u8);

                int col = rRef.IsEmpty ? _nextCol : XlsxXml.ColumnIndex(rRef);
                if (col < 0)
                {
                    col = _nextCol;
                }
                _nextCol = col + 1;
                int style = ParseInt(sVal);
                var kind = ClassifyKind(tVal);
                bool selfClose = _buf[gt - 1] == '/';
                _pos = gt + 1; // consume open tag; _pos now at inner start (or next element if self-closed)
                return new CellHeader(col, style, kind, selfClose);
            }

            private enum Kind { Number, Shared, Inline, Bool, Error, Formula }

            // "" / "n" -> Number; "s" shared; "inlineStr" inline; "b" bool; "e" error; "str" formula result.
            private static Kind ClassifyKind(ReadOnlySpan<byte> t)
            {
                if (t.SequenceEqual("s"u8))
                {
                    return Kind.Shared;
                }
                if (t.SequenceEqual("inlineStr"u8))
                {
                    return Kind.Inline;
                }
                if (t.SequenceEqual("b"u8))
                {
                    return Kind.Bool;
                }
                if (t.SequenceEqual("e"u8))
                {
                    return Kind.Error;
                }
                if (t.SequenceEqual("str"u8))
                {
                    return Kind.Formula;
                }

                return Kind.Number;
            }


            private void EmitCell(Kind kind, ReadOnlySpan<byte> inner, int col, int style)
            {
                // Shared strings: <v> holds an index; point the cell at that slice of the shared buffer.
                if (kind == Kind.Shared)
                {
                    var (start, len) = _reader.SharedAt(ParseInt(ElementText(inner, "<v>"u8, "</v>"u8)));
                    AddCell(col, start, len, CellType.ExcelString, style, fromShared: true);
                    return;
                }

                int vStart = _valLen;
                if (kind == Kind.Inline)
                {
                    EnsureValsCapacity(_valLen + inner.Length);
                    _valLen += XlsxXml.WriteTextRuns(inner, _vals.AsSpan(_valLen));
                    AddCell(col, vStart, _valLen - vStart, CellType.ExcelString, style, fromShared: false);
                    return;
                }

                CellType cellType = kind switch
                {
                    Kind.Bool => CellType.Boolean,
                    Kind.Error => CellType.Error,
                    Kind.Formula => CellType.Formula,
                    _ => _reader.IsDateStyle(style) ? CellType.Date : CellType.Number,
                };
                AppendDecoded(ElementText(inner, "<v>"u8, "</v>"u8));
                AddCell(col, vStart, _valLen - vStart, cellType, style, fromShared: false);
            }

            private static ReadOnlySpan<byte> ElementText(ReadOnlySpan<byte> inner, ReadOnlySpan<byte> openTag, ReadOnlySpan<byte> closeTag)
            {
                int s = inner.IndexOf(openTag);
                if (s < 0)
                {
                    return default;
                }
                s += openTag.Length;
                int e = inner[s..].IndexOf(closeTag);
                return e < 0 ? default : inner.Slice(s, e);
            }

            private void AppendDecoded(ReadOnlySpan<byte> src)
            {
                if (src.IsEmpty)
                {
                    return;
                }
                EnsureValsCapacity(_valLen + src.Length);
                _valLen += XlsxXml.Decode(src, _vals.AsSpan(_valLen));
            }

            private void AddCell(int col, int start, int len, CellType type, int style, bool fromShared)
            {
                if (_cellCount == _cells.Length)
                {
                    var bigger = ArrayPool<CellDesc>.Shared.Rent(_cells.Length * 2);
                    Array.Copy(_cells, bigger, _cellCount);
                    ArrayPool<CellDesc>.Shared.Return(_cells);
                    _cells = bigger;
                }
                _cells[_cellCount++] = new CellDesc
                {
                    Column = col,
                    Start = start,
                    Length = len,
                    Type = type,
                    Style = style,
                    FromShared = fromShared,
                };
            }

            private void EnsureValsCapacity(int needed)
            {
                if (needed <= _vals.Length)
                {
                    return;
                }
                var bigger = ArrayPool<byte>.Shared.Rent(Math.Max(_vals.Length * 2, needed));
                Array.Copy(_vals, bigger, _valLen);
                ArrayPool<byte>.Shared.Return(_vals);
                _vals = bigger;
            }

            private static int ParseInt(ReadOnlySpan<byte> src)
            {
                return Utf8Parser.TryParse(src, out int v, out _) ? v : 0;
            }

            private static bool IsBoundary(byte b)
            {
                return b is (byte)' ' or (byte)'>' or (byte)'/' or (byte)'\t' or (byte)'\r' or (byte)'\n';
            }

            // --- buffer management ---
            // After every Fill the window [_pos.._len) is rescanned from the start; that re-reads a few
            // bytes but keeps the search loops trivial and handles delimiters split across a refill for free.

            private int IndexOf(byte b)
            {
                while (true)
                {
                    int rel = _buf.AsSpan(_pos, _len - _pos).IndexOf(b);
                    if (rel >= 0)
                    {
                        return _pos + rel;
                    }
                    if (_eof)
                    {
                        return -1;
                    }
                    Fill();
                }
            }

            private int IndexOfSeq(ReadOnlySpan<byte> seq)
            {
                while (true)
                {
                    int rel = _buf.AsSpan(_pos, _len - _pos).IndexOf(seq);
                    if (rel >= 0)
                    {
                        return _pos + rel;
                    }
                    if (_eof)
                    {
                        return -1;
                    }
                    Fill();
                }
            }

            private void Ensure(int n)
            {
                while (_len - _pos < n && !_eof)
                {
                    Fill();
                }
            }

            // Make room for more bytes (compact consumed prefix, else grow), then read once.
            private void Fill()
            {
                if (_pos > 0)
                {
                    _buf.AsSpan(_pos, _len - _pos).CopyTo(_buf);
                    _len -= _pos;
                    _pos = 0;
                }
                else if (_len == _buf.Length)
                {
                    var bigger = ArrayPool<byte>.Shared.Rent(_buf.Length * 2);
                    _buf.AsSpan(0, _len).CopyTo(bigger);
                    ArrayPool<byte>.Shared.Return(_buf);
                    _buf = bigger;
                }
                int n = _sheet!.Read(_buf, _len, _buf.Length - _len);
                if (n == 0)
                {
                    _eof = true;
                }
                else
                {
                    _len += n;
                }
            }

            // Async twins of the search/refill primitives. Each searches the buffered window first and
            // only awaits a refill on a miss, so no span is ever held across the await.

            private async ValueTask<int> IndexOfAsync(byte b)
            {
                while (true)
                {
                    int rel = _buf.AsSpan(_pos, _len - _pos).IndexOf(b);
                    if (rel >= 0)
                    {
                        return _pos + rel;
                    }
                    if (_eof)
                    {
                        return -1;
                    }
                    await FillAsync().ConfigureAwait(false);
                }
            }

            // Only one multi-byte terminator is searched on the async path, so it's hardcoded rather than
            // taking a ReadOnlySpan<byte> (which can't be an async method parameter).
            private async ValueTask<int> IndexOfCloseCellAsync()
            {
                while (true)
                {
                    int rel = _buf.AsSpan(_pos, _len - _pos).IndexOf("</c>"u8);
                    if (rel >= 0)
                    {
                        return _pos + rel;
                    }
                    if (_eof)
                    {
                        return -1;
                    }
                    await FillAsync().ConfigureAwait(false);
                }
            }

            private async ValueTask EnsureAsync(int n)
            {
                while (_len - _pos < n && !_eof)
                {
                    await FillAsync().ConfigureAwait(false);
                }
            }

            private async ValueTask FillAsync()
            {
                if (_pos > 0)
                {
                    _buf.AsSpan(_pos, _len - _pos).CopyTo(_buf);
                    _len -= _pos;
                    _pos = 0;
                }
                else if (_len == _buf.Length)
                {
                    var bigger = ArrayPool<byte>.Shared.Rent(_buf.Length * 2);
                    _buf.AsSpan(0, _len).CopyTo(bigger);
                    ArrayPool<byte>.Shared.Return(_buf);
                    _buf = bigger;
                }
                int n = await _sheet!.ReadAsync(_buf.AsMemory(_len, _buf.Length - _len), _ct).ConfigureAwait(false);
                if (n == 0)
                {
                    _eof = true;
                }
                else
                {
                    _len += n;
                }
            }

            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
                Justification = "_sheet is opened for this enumerator and owned by it.")]
            public void Dispose()
            {
                _sheet?.Dispose();
                _sheet = null;
                ReturnBuffers();
            }

            [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
                Justification = "_sheet is opened for this enumerator and owned by it.")]
            public async ValueTask DisposeAsync()
            {
                if (_sheet is not null)
                {
                    await _sheet.DisposeAsync().ConfigureAwait(false);
                    _sheet = null;
                }
                ReturnBuffers();
            }

            private void ReturnBuffers()
            {
                if (_buf.Length > 0)
                {
                    ArrayPool<byte>.Shared.Return(_buf);
                    _buf = [];
                }
                if (_vals.Length > 0)
                {
                    ArrayPool<byte>.Shared.Return(_vals);
                    _vals = [];
                }
                if (_cells.Length > 0)
                {
                    ArrayPool<CellDesc>.Shared.Return(_cells);
                    _cells = [];
                }
            }
        }
    }
}
