using System.Buffers;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using static ExcelReader.Core.Reader.Biff12;

namespace ExcelReader.Core.Reader
{
    /// <summary>Reads rows from a legacy BIFF8 (.xls) workbook, exposing each worksheet through synchronous and asynchronous enumeration.</summary>
    public sealed partial class XlsReader : IExcelRowReader, IExcelRowReader<XlsReader.Enumerator>
    {
        private readonly WorkbookStream _workbook;
        private readonly ExcelReaderOptions _options;
        private readonly (string Name, int Offset)[] _sheets;
        private readonly bool[] _styleIsDate;
        private readonly bool _date1904;
        private byte[] _sharedFlat;
        private int[] _sharedOffsets;
        // Lazily created: dedups repeated LABELSST values (categorical columns) into one string
        // instance instead of re-decoding UTF-8 per row. Keyed by the string's stable byte offset into
        // _sharedFlat (see CellDesc.ToCell / Cell.GetString) — same shape as XlsxReader/XlsbReader.
        private Dictionary<int, string>? _sharedStringCache;
        private int _current;

        internal XlsReader(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null)
            : this(XlsCompoundFile.OpenWorkbook(stream, leaveOpen), options)
        {
        }

        internal XlsReader(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
            : this(XlsCompoundFile.OpenWorkbook(data), options)
        {
        }

        private XlsReader(WorkbookStream workbook, ExcelReaderOptions? options = null)
        {
            _workbook = workbook;
            _options = options ?? ExcelReaderOptions.Default;
            using (BiffCursor cursor = workbook.OpenCursor())
            {
                ParseWorkbookGlobals(cursor, _options, out _sheets, out _styleIsDate, out _date1904, out _sharedFlat, out _sharedOffsets);
            }
            if (_sheets.Length == 0)
            {
#pragma warning disable IDISP007 // Ownership of the WorkbookStream transferred to this reader; dispose it on the no-sheets failure path.
                workbook.Dispose();
#pragma warning restore IDISP007
                throw new InvalidDataException("The workbook contains no sheets.");
            }
        }

        internal static async ValueTask<XlsReader> CreateAsync(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null, CancellationToken ct = default)
        {
            WorkbookStream workbook = await XlsCompoundFile.OpenWorkbookAsync(stream, leaveOpen, ct).ConfigureAwait(false);
            return new XlsReader(workbook, options);
        }

        internal BiffCursor OpenCursor(int offset)
        {
            BiffCursor cursor = _workbook.OpenCursor();
            cursor.Position = offset;
            return cursor;
        }

        /// <inheritdoc/>
        public string SheetName => _sheets[_current].Name;

        /// <inheritdoc/>
        public int SheetCount => _sheets.Length;

        /// <inheritdoc/>
        public bool IsDate1904 => _date1904;

        internal ReadOnlySpan<byte> SharedSpan => _sharedFlat;

        internal Dictionary<int, string> SharedStringCache => _sharedStringCache ??= [];

        internal bool IsDateStyle(int style)
        {
            return WorkbookLookups.IsDateStyle(_styleIsDate, style);
        }

        internal (int Start, int Length) SharedAt(int index)
        {
            return WorkbookLookups.SharedAt(_sharedOffsets, index);
        }

        /// <inheritdoc/>
        public bool TryMoveToSheet(ReadOnlySpan<char> name)
        {
            if (!WorkbookLookups.TryFindSheetIndex(_sheets, name, static s => s.Name, out int index))
            {
                return false;
            }
            _current = index;
            return true;
        }

        /// <inheritdoc/>
        public void MoveToSheet(int index)
        {
            WorkbookLookups.ValidateSheetIndex(index, _sheets.Length);
            _current = index;
        }

        /// <inheritdoc/>
        [SuppressMessage("Performance", "HLQ006:GetEnumerator should return a value type",
            Justification = "Enumerator is a class so the same type can also expose MoveNextAsync for parity with XlsxReader.")]
        public Enumerator GetEnumerator()
        {
            return new Enumerator(this, _sheets[_current].Offset);
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <inheritdoc/>
        [SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "Enumerator is a class so the same type can also expose MoveNextAsync for parity with XlsxReader.")]
        public Enumerator GetAsyncEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>Creates an enumerator for the current sheet that observes cancellation while iterating.</summary>
        /// <remarks>This overload takes no default value for <paramref name="ct"/>: the parameterless <see cref="GetAsyncEnumerator()"/> above already covers the no-argument call, so a default here would only shadow it.</remarks>
        /// <param name="ct">Token checked before enumeration starts and on each subsequent move.</param>
        [SuppressMessage("Performance", "HLQ006:GetAsyncEnumerator should return a value type",
            Justification = "Enumerator is a class so the same type can also expose MoveNextAsync for parity with XlsxReader.")]
        public Enumerator GetAsyncEnumerator(CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            return new Enumerator(this, _sheets[_current].Offset, ct);
        }

        IExcelRowEnumerator IExcelRowReader<IExcelRowEnumerator>.GetAsyncEnumerator()
        {
            return GetAsyncEnumerator();
        }

        /// <inheritdoc/>
        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Enumerator ownership transfers to the caller, who disposes it via await using / DisposeAsync.")]
        public ValueTask<Enumerator> GetAsyncEnumeratorAsync(CancellationToken ct = default)
        {
            return new ValueTask<Enumerator>(GetAsyncEnumerator(ct));
        }

        [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "Enumerator ownership transfers to the caller, who disposes it via await using / DisposeAsync.")]
        [SuppressMessage("Performance", "CA1849:Call async methods when in an async method",
            Justification = "XlsReader is fully in-memory; opening the enumerator is synchronous, so there is no async open to await here.")]
        ValueTask<IExcelRowEnumerator> IExcelRowReader<IExcelRowEnumerator>.GetAsyncEnumeratorAsync(CancellationToken ct)
        {
            return new ValueTask<IExcelRowEnumerator>(GetAsyncEnumerator(ct));
        }

        /// <inheritdoc/>
        public void Dispose()
        {
            _workbook.Dispose();
            _sharedFlat = [];
            _sharedOffsets = [0];
            _sharedStringCache = null;
        }

        /// <inheritdoc/>
        public ValueTask DisposeAsync()
        {
            Dispose();
            return ValueTask.CompletedTask;
        }

        private static void ParseWorkbookGlobals(
            BiffCursor cursor,
            ExcelReaderOptions options,
            out (string Name, int Offset)[] sheets,
            out bool[] styleIsDate,
            out bool date1904,
            out byte[] sharedFlat,
            out int[] sharedOffsets)
        {
            List<(string Name, int Offset)> sheetList = [];
            Dictionary<int, bool> customFormats = new(capacity: 16);
            List<bool> styleFlags = [];
            date1904 = false;
            sharedFlat = [];
            sharedOffsets = [0];

            bool sawGlobalsBof = false;
            while (cursor.TryReadRecord(out int id, out ReadOnlySpan<byte> data))
            {
                if (id == Rec.Bof)
                {
                    if (data.Length < 4 || ReadU16(data, 0) != Biff8Version)
                    {
                        throw new NotSupportedException("Only BIFF8 .xls workbooks are supported.");
                    }
                    sawGlobalsBof = ReadU16(data, 2) == SubstreamGlobals;
                    continue;
                }
                if (id == Rec.Eof && sawGlobalsBof)
                {
                    break;
                }
                switch (id)
                {
                    case Rec.Date1904:
                        date1904 = data.Length >= 2 && ReadU16(data, 0) != 0;
                        break;
                    case Rec.BoundSheet:
                        if (TryParseBoundSheet(data, out var sheet))
                        {
                            sheetList.Add(sheet);
                        }
                        break;
                    case Rec.Sst:
                        DecodeSstFromCursor(cursor, data, options, out sharedFlat, out sharedOffsets);
                        break;
                    case Rec.Format:
                        if (TryParseFormat(data, out int formatId, out string format))
                        {
                            customFormats[formatId] = NumberFormat.LooksLikeDate(format);
                        }
                        break;
                    case Rec.Xf:
                        if (data.Length >= 4)
                        {
                            int formatIndex = ReadU16(data, 2);
                            styleFlags.Add(WorkbookLookups.ResolveDateFlag(customFormats, formatIndex));
                        }
                        break;
                    case Rec.FilePass:
                        throw new NotSupportedException("Encrypted .xls workbooks are not supported.");
                }
            }

            sheets = [.. sheetList];
            styleIsDate = [.. styleFlags];
        }

        private static bool TryParseBoundSheet(ReadOnlySpan<byte> data, out (string Name, int Offset) sheet)
        {
            sheet = default;
            // BoundSheet8 byte 5 is the sheet type (0 = worksheet). Charts, macro sheets and
            // dialog sheets have a different substream layout and must not be enumerated as rows.
            if (data.Length < 8 || data[5] != 0
                || !TryDecodeBiffString(data, start: 8, charCount: data[6], flags: data[7], out string name))
            {
                return false;
            }
            sheet = (name, ReadI32(data, 0));
            return true;
        }

        private static bool TryParseFormat(ReadOnlySpan<byte> data, out int formatId, out string format)
        {
            formatId = 0;
            format = string.Empty;
            if (data.Length < 5 || !TryDecodeBiffString(data, start: 5, charCount: ReadU16(data, 2), flags: data[4], out format))
            {
                return false;
            }
            formatId = ReadU16(data, 0);
            return true;
        }

        // BIFF8 string: bit 0 of the flags byte picks compressed (1 byte/char) vs UTF-16.
        private static bool TryDecodeBiffString(ReadOnlySpan<byte> data, int start, int charCount, byte flags, out string value)
        {
            bool compressed = (flags & 1) == 0;
            int byteCount = compressed ? charCount : charCount * 2;
            if (start + byteCount > data.Length)
            {
                value = string.Empty;
                return false;
            }
            ReadOnlySpan<byte> raw = data.Slice(start, byteCount);
            value = compressed
                ? DecodeCompressedString(raw, charCount)
                : System.Text.Encoding.Unicode.GetString(raw);
            return true;
        }

        // The SST payload (first record minus its 8-byte header) plus any CONTINUE records must be
        // contiguous because strings can straddle record boundaries. Gather into a pooled buffer
        // off the cursor, then decode into the retained flat buffer.
        private static void DecodeSstFromCursor(
            BiffCursor cursor,
            ReadOnlySpan<byte> first,
            ExcelReaderOptions options,
            out byte[] sharedFlat,
            out int[] sharedOffsets)
        {
            int initialLen = first.Length > 8 ? first.Length - 8 : 0;
            LimitChecks.ThrowIfOverSharedStringLimit(options, initialLen);
            byte[] buffer = ArrayPool<byte>.Shared.Rent(Math.Max(256, initialLen));
            int len = 0;
            if (first.Length > 8)
            {
                first[8..].CopyTo(buffer);
                len = initialLen;
            }
            // A string's character array can be split across a CONTINUE boundary; when it is, the record
            // resumes with a fresh grbit (compression) byte that is NOT part of the character data
            // ([MS-XLS] 2.5.240). Record where each CONTINUE payload begins so the decoder can consume
            // that byte instead of misreading it as a character (which corrupts every following string).
            List<int> boundaries = [];
            while (cursor.PeekId() == Rec.Continue && cursor.TryReadRecord(out _, out ReadOnlySpan<byte> cont))
            {
                boundaries.Add(len);
                EnsureSharedCapacity(options, ref buffer, len + cont.Length);
                cont.CopyTo(buffer.AsSpan(len));
                len += cont.Length;
            }
            DecodeSharedStrings(buffer.AsSpan(0, len), CollectionsMarshal.AsSpan(boundaries), options, out sharedFlat, out sharedOffsets);
            ArrayPool<byte>.Shared.Return(buffer);
        }

        // `boundaries` holds the offsets (into `sst`) where each CONTINUE payload starts. A boundary that
        // falls inside a string's character array marks an inserted grbit byte to consume; one outside the
        // array (header, formatting runs, extended data) carries no grbit and is simply read through. A
        // wide (UTF-16) code unit whose two bytes straddle the boundary is handled too: its low byte is
        // read from the end of the preceding run, then the grbit is consumed, then its high byte from the
        // start of the continuation — see the `boundaries[boundaryIdx] == pos + 1` branch below.
        private static void DecodeSharedStrings(ReadOnlySpan<byte> sst, ReadOnlySpan<int> boundaries, ExcelReaderOptions options, out byte[] sharedFlat, out int[] sharedOffsets)
        {
            LimitChecks.ThrowIfOverSharedStringLimit(options, sst.Length);
            byte[] flat = ArrayPool<byte>.Shared.Rent(Math.Max(256, sst.Length * 3));
            // One string's decoded UTF-16 units; each unit consumes >= 1 source byte, so sst.Length caps it.
            char[] scratch = ArrayPool<char>.Shared.Rent(Math.Max(64, sst.Length));
            int flatLen = 0;
            List<int> offsets = [0];
            int pos = 0;
            int boundaryIdx = 0;
            try
            {
                while (pos + 3 <= sst.Length)
                {
                    int chars = ReadU16(sst, pos);
                    byte flags = sst[pos + 2];
                    pos += 3;
                    int richRuns = 0;
                    int extBytes = 0;
                    if ((flags & 0x08) != 0)
                    {
                        if (pos + 2 > sst.Length) { break; }
                        richRuns = ReadU16(sst, pos);
                        pos += 2;
                    }
                    if ((flags & 0x04) != 0)
                    {
                        if (pos + 4 > sst.Length) { break; }
                        extBytes = ReadI32(sst, pos);
                        pos += 4;
                    }
                    bool compressed = (flags & 1) == 0;
                    int produced = DecodeChars(sst, boundaries, chars, scratch, ref pos, ref boundaryIdx, ref compressed, out bool truncated);
                    int maxBytes = System.Text.Encoding.UTF8.GetMaxByteCount(produced);
                    EnsureSharedCapacity(options, ref flat, flatLen + maxBytes);
                    flatLen += System.Text.Encoding.UTF8.GetBytes(scratch.AsSpan(0, produced), flat.AsSpan(flatLen));
                    offsets.Add(flatLen);
                    if (truncated) { break; }
                    // Formatting runs (4 bytes each) and extended data follow the characters, split across
                    // boundaries without a grbit — skip them straight through, bailing on bogus lengths.
                    long next = pos + ((long)richRuns * 4) + extBytes;
                    if (richRuns < 0 || extBytes < 0 || next > sst.Length) { break; }
                    pos = (int)next;
                }

                sharedFlat = flat[..flatLen];
                sharedOffsets = [.. offsets];
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(flat);
                ArrayPool<char>.Shared.Return(scratch);
            }
        }

        // Decodes one shared string's code units into scratch, advancing the position and boundary
        // index and flipping the compression mode at each grbit boundary. Returns the unit count
        // produced, and reports truncated when the source ends mid-string.
        private static int DecodeChars(ReadOnlySpan<byte> sst, ReadOnlySpan<int> boundaries, int chars, char[] scratch, ref int pos, ref int boundaryIdx, ref bool compressed, out bool truncated)
        {
            truncated = false;
            int produced = 0;
            for (int c = 0; c < chars; c++)
            {
                // Drop boundaries already behind us (splits outside the character array), then
                // consume the grbit for a boundary that lands exactly on this character.
                while (boundaryIdx < boundaries.Length && boundaries[boundaryIdx] < pos) { boundaryIdx++; }
                if (boundaryIdx < boundaries.Length && boundaries[boundaryIdx] == pos)
                {
                    if (pos >= sst.Length) { truncated = true; break; }
                    compressed = (sst[pos] & 1) == 0;
                    pos++;
                    boundaryIdx++;
                }

                // A UTF-16 code unit may itself straddle a CONTINUE boundary. Its low byte belongs to
                // the preceding run, followed by the continuation grbit, followed by its high byte.
                // Decode that unit as wide regardless of the new grbit; the grbit controls the next unit.
                if (!compressed && boundaryIdx < boundaries.Length && boundaries[boundaryIdx] == pos + 1)
                {
                    if (pos + 2 >= sst.Length) { truncated = true; break; }
                    byte low = sst[pos++];
                    compressed = (sst[pos++] & 1) == 0;
                    boundaryIdx++;
                    scratch[produced++] = (char)(low | (sst[pos++] << 8));
                    continue;
                }
                int step = compressed ? 1 : 2;
                if (pos + step > sst.Length) { truncated = true; break; }
                scratch[produced++] = compressed
                    ? DecodeCp1252(sst[pos])
                    : (char)(sst[pos] | (sst[pos + 1] << 8));
                pos += step;
            }
            return produced;
        }

        private static void EnsureSharedCapacity(ExcelReaderOptions options, ref byte[] buffer, int needed)
        {
            if (needed <= buffer.Length)
            {
                return;
            }
            LimitChecks.ThrowIfOverSharedStringLimit(options, needed);
            byte[] bigger = ArrayPool<byte>.Shared.Rent(Math.Max(buffer.Length * 2, needed));
            buffer.CopyTo(bigger, 0);
            ArrayPool<byte>.Shared.Return(buffer);
            buffer = bigger;
        }

        private const int Biff8Version = 0x0600;
        private const int SubstreamGlobals = 0x0005;
        private const int SubstreamWorksheet = 0x0010;

        // BIFF8 record type IDs (see [MS-XLS]).
        private static class Rec
        {
            internal const int Bof = 0x0809;
            internal const int Eof = 0x000A;
            internal const int Date1904 = 0x0022;
            internal const int BoundSheet = 0x0085;
            internal const int Sst = 0x00FC;
            internal const int Continue = 0x003C;
            internal const int Format = 0x041E;
            internal const int Xf = 0x00E0;
            internal const int FilePass = 0x002F;
            internal const int Label = 0x0204;
            internal const int LabelSst = 0x00FD;
            internal const int Number = 0x0203;
            internal const int Rk = 0x027E;
            internal const int MulRk = 0x00BD;
            internal const int BoolErr = 0x0205;
            internal const int Formula = 0x0006;
            internal const int Blank = 0x0201;
            internal const int MulBlank = 0x00BE;
            internal const int StringRec = 0x0207;
        }
    }
}
