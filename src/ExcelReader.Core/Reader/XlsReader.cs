using System.Buffers;
using System.Buffers.Binary;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace ExcelReader.Core.Reader
{
    public sealed partial class XlsReader : IDisposable, IAsyncDisposable
    {
        private readonly byte[] _workbook;
        private readonly (string Name, int Offset)[] _sheets;
        private readonly bool[] _styleIsDate;
        private readonly bool _date1904;
        private byte[] _sharedFlat;
        private int[] _sharedOffsets;
        private int _current;

        internal XlsReader(Stream stream, bool leaveOpen)
            : this(XlsCompoundFile.Open(stream, leaveOpen).ReadWorkbookStream())
        {
        }

        private XlsReader(byte[] workbook)
        {
            _workbook = workbook;
            ParseWorkbookGlobals(workbook, out _sheets, out _styleIsDate, out _date1904, out _sharedFlat, out _sharedOffsets);
            if (_sheets.Length == 0)
            {
                throw new InvalidDataException("The workbook contains no sheets.");
            }
        }

        internal static async ValueTask<XlsReader> CreateAsync(Stream stream, bool leaveOpen, CancellationToken ct)
        {
            XlsCompoundFile ole = await XlsCompoundFile.OpenAsync(stream, leaveOpen, ct).ConfigureAwait(false);
            return new XlsReader(ole.ReadWorkbookStream());
        }

        public string SheetName => _sheets[_current].Name;
        public int SheetCount => _sheets.Length;
        public bool IsDate1904 => _date1904;

        internal ReadOnlySpan<byte> SharedSpan => _sharedFlat;

        internal bool IsDateStyle(int style)
        {
            return (uint)style < (uint)_styleIsDate.Length && _styleIsDate[style];
        }

        internal (int Start, int Length) SharedAt(int index)
        {
            if ((uint)index >= (uint)(_sharedOffsets.Length - 1))
            {
                return (0, 0);
            }
            return (_sharedOffsets[index], _sharedOffsets[index + 1] - _sharedOffsets[index]);
        }

        public bool TryMoveToSheet(ReadOnlySpan<char> name)
        {
            for (int i = 0; i < _sheets.Length; i++)
            {
                if (name.Equals(_sheets[i].Name, StringComparison.OrdinalIgnoreCase))
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
            Justification = "Enumerator is a class so the same type can also expose MoveNextAsync for parity with XlsxReader.")]
        public Enumerator GetEnumerator()
        {
            return new Enumerator(this, _sheets[_current].Offset);
        }

        public ValueTask<Enumerator> GetAsyncEnumeratorAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            return new ValueTask<Enumerator>(new Enumerator(this, _sheets[_current].Offset, ct));
        }

        public void Dispose()
        {
            _sharedFlat = [];
            _sharedOffsets = [0];
        }

        public ValueTask DisposeAsync()
        {
            Dispose();
            return ValueTask.CompletedTask;
        }

        private static void ParseWorkbookGlobals(
            ReadOnlySpan<byte> workbook,
            out (string Name, int Offset)[] sheets,
            out bool[] styleIsDate,
            out bool date1904,
            out byte[] sharedFlat,
            out int[] sharedOffsets)
        {
            List<(string Name, int Offset)> sheetList = [];
            Dictionary<int, bool> customFormats = new(capacity: 16);
            List<bool> styleFlags = [];
            List<SstSpan> sstSpans = [];
            date1904 = false;
            sharedFlat = [];
            sharedOffsets = [0];

            int pos = 0;
            bool sawGlobalsBof = false;
            while (TryReadRecord(workbook, ref pos, out int id, out ReadOnlySpan<byte> data))
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
                        // Store offsets into the retained workbook buffer instead of copying
                        // each record; the 8-byte SST header is folded into the start offset.
                        sstSpans.Add(new SstSpan(pos - data.Length + 8, data.Length - 8));
                        while (PeekRecordId(workbook, pos) == Rec.Continue && TryReadRecord(workbook, ref pos, out _, out ReadOnlySpan<byte> cont))
                        {
                            sstSpans.Add(new SstSpan(pos - cont.Length, cont.Length));
                        }
                        ParseSharedStrings(workbook, sstSpans, out sharedFlat, out sharedOffsets);
                        break;
                    case Rec.Format:
                        if (TryParseFormat(data, out int formatId, out string format))
                        {
                            customFormats[formatId] = LooksLikeDateFormat(format);
                        }
                        break;
                    case Rec.Xf:
                        if (data.Length >= 4)
                        {
                            int formatIndex = ReadU16(data, 2);
                            styleFlags.Add(customFormats.TryGetValue(formatIndex, out bool custom)
                                ? custom
                                : IsBuiltinDateFormat(formatIndex));
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
            if (data.Length < 8 || !TryDecodeBiffString(data, start: 8, charCount: data[6], flags: data[7], out string name))
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

        private static void ParseSharedStrings(ReadOnlySpan<byte> workbook, List<SstSpan> spans, out byte[] sharedFlat, out int[] sharedOffsets)
        {
            // Single record (the common case): decode straight from the workbook, no concat copy.
            if (spans.Count == 1)
            {
                DecodeSharedStrings(workbook.Slice(spans[0].Start, spans[0].Length), out sharedFlat, out sharedOffsets);
                return;
            }

            // Multiple records: strings may straddle CONTINUE boundaries, so the record
            // payloads must be contiguous. Use a pooled scratch buffer for the concat.
            int total = 0;
            for (int i = 0; i < spans.Count; i++)
            {
                total += spans[i].Length;
            }
            byte[] sst = ArrayPool<byte>.Shared.Rent(total);
            int written = 0;
            for (int i = 0; i < spans.Count; i++)
            {
                workbook.Slice(spans[i].Start, spans[i].Length).CopyTo(sst.AsSpan(written));
                written += spans[i].Length;
            }
            DecodeSharedStrings(sst.AsSpan(0, total), out sharedFlat, out sharedOffsets);
            ArrayPool<byte>.Shared.Return(sst);
        }

        private static void DecodeSharedStrings(ReadOnlySpan<byte> sst, out byte[] sharedFlat, out int[] sharedOffsets)
        {
            byte[] flat = ArrayPool<byte>.Shared.Rent(Math.Max(256, sst.Length * 3));
            int flatLen = 0;
            List<int> offsets = [0];
            int pos = 0;
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
                int needed = (flags & 1) == 0 ? chars : chars * 2;
                if (pos + needed > sst.Length)
                {
                    break;
                }
                EnsureCapacity(ref flat, flatLen + (chars * 4));
                flatLen += DecodeStringToUtf8(sst.Slice(pos, needed), chars, flags, flat.AsSpan(flatLen));
                pos += needed + (richRuns * 4) + extBytes;
                offsets.Add(flatLen);
            }

            sharedFlat = flat.AsSpan(0, flatLen).ToArray();
            ArrayPool<byte>.Shared.Return(flat);
            sharedOffsets = [.. offsets];
        }

        private static bool TryReadRecord(ReadOnlySpan<byte> src, ref int pos, out int id, out ReadOnlySpan<byte> data)
        {
            id = 0;
            data = default;
            if (pos + 4 > src.Length)
            {
                return false;
            }
            id = ReadU16(src, pos);
            int len = ReadU16(src, pos + 2);
            pos += 4;
            if (pos + len > src.Length)
            {
                return false;
            }
            data = src.Slice(pos, len);
            pos += len;
            return true;
        }

        private static int PeekRecordId(ReadOnlySpan<byte> src, int pos)
        {
            return pos + 4 <= src.Length ? ReadU16(src, pos) : -1;
        }

        private static ushort ReadU16(ReadOnlySpan<byte> src, int offset)
        {
            return BinaryPrimitives.ReadUInt16LittleEndian(src.Slice(offset, 2));
        }

        private static int ReadI32(ReadOnlySpan<byte> src, int offset)
        {
            return BinaryPrimitives.ReadInt32LittleEndian(src.Slice(offset, 4));
        }

        private static uint ReadU32(ReadOnlySpan<byte> src, int offset)
        {
            return BinaryPrimitives.ReadUInt32LittleEndian(src.Slice(offset, 4));
        }

        private static void EnsureCapacity(ref byte[] buffer, int needed)
        {
            if (needed <= buffer.Length)
            {
                return;
            }
            byte[] bigger = ArrayPool<byte>.Shared.Rent(Math.Max(buffer.Length * 2, needed));
            buffer.CopyTo(bigger, 0);
            ArrayPool<byte>.Shared.Return(buffer);
            buffer = bigger;
        }
        [StructLayout(LayoutKind.Auto)]
        private readonly record struct SstSpan(int Start, int Length);

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
        }
    }
}
