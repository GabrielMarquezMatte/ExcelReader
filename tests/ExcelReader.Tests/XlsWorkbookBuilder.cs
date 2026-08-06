using System.Buffers.Binary;
using System.Text;
using System.Runtime.InteropServices;

namespace ExcelReader.Tests
{
    internal sealed record XlsSharedString(string Value);
    internal sealed record XlsSharedIndex(int Index);
    internal sealed record XlsUnicodeString(string Value);
    internal sealed record XlsCompressedBytes(byte[] Bytes);
    internal sealed record XlsDate(DateTime Value);
    internal sealed record XlsError(byte Code);
    internal sealed record XlsFormula(double Value);
    internal sealed record XlsFormulaBool(bool Value);
    internal sealed record XlsFormulaError(byte Code);
    internal sealed record XlsRkInt(int Value);
    internal sealed record XlsRkRaw(uint Value);
    internal sealed record XlsMulRk(params int[] Values);
    internal sealed record XlsBlank;
    internal sealed record XlsMulBlank(int Count);
    internal sealed record XlsAt(int Column, object? Value);

    internal static class XlsWorkbookBuilder
    {
        private const int SectorSize = 512;
        private const int FatSector = unchecked((int)0xFFFFFFFD);
        private const int EndOfChain = unchecked((int)0xFFFFFFFE);
        private const int FreeSector = unchecked((int)0xFFFFFFFF);

        internal static MemoryStream Build(bool date1904 = false, string? customDateFormat = null, params (string Name, object?[][] Rows)[] sheets)
        {
            List<string> sharedStrings = [];
            Dictionary<string, int> sharedIndexes = new(StringComparer.Ordinal);
            byte[][] sheetStreams = new byte[sheets.Length][];
            for (int i = 0; i < sheets.Length; i++)
            {
                sheetStreams[i] = BuildSheet(sheets[i].Rows, sharedStrings, sharedIndexes, date1904);
            }

            byte[] globals = BuildGlobals(sheets, sheetStreams, sharedStrings, date1904, customDateFormat);
            using MemoryStream workbook = new();
            workbook.Write(globals);
            foreach (byte[] sheet in sheetStreams)
            {
                workbook.Write(sheet);
            }

            return BuildOle(workbook.ToArray());
        }

        internal static MemoryStream BuildNonBiff8()
        {
            byte[] globals = Record(0x0809, [.. U16(0x0500), .. U16(0x0005), .. new byte[12]]);
            globals = [.. globals, .. Record(0x000A, [])];
            return BuildOle(globals);
        }

        internal static MemoryStream BuildEncrypted()
        {
            using MemoryStream globals = new();
            WriteRecord(globals, 0x0809, [.. U16(0x0600), .. U16(0x0005), .. new byte[12]]);
            WriteRecord(globals, 0x002F, [0, 0]);
            WriteRecord(globals, 0x000A, []);
            return BuildOle(globals.ToArray());
        }

        internal static MemoryStream BuildNoSheets()
        {
            using MemoryStream globals = new();
            WriteRecord(globals, 0x0809, [.. U16(0x0600), .. U16(0x0005), .. new byte[12]]);
            WriteRecord(globals, 0x000A, []);
            return BuildOle(globals.ToArray());
        }

        internal static MemoryStream BuildBadSheetBof()
        {
            byte[] badSheet = Record(0x0809, [.. U16(0x0500), .. U16(0x0010), .. new byte[12]]);
            const int globalsLength = 20 + 14 + 4;
            using MemoryStream workbook = new();
            WriteRecord(workbook, 0x0809, [.. U16(0x0600), .. U16(0x0005), .. new byte[12]]);
            WriteRecord(workbook, 0x0085, BoundSheet(globalsLength, "S1"));
            WriteRecord(workbook, 0x000A, []);
            workbook.Write(badSheet);
            return BuildOle(workbook.ToArray());
        }

        internal static MemoryStream BuildRawSheet(bool includeEof, params (int Id, byte[] Data)[] records)
        {
            using MemoryStream sheet = new();
            WriteRecord(sheet, 0x0809, [.. U16(0x0600), .. U16(0x0010), .. new byte[12]]);
            foreach ((int id, byte[] data) in records)
            {
                WriteRecord(sheet, id, data);
            }
            if (includeEof)
            {
                WriteRecord(sheet, 0x000A, []);
            }

            byte[] sheetBytes = sheet.ToArray();
            const int globalsLength = 20 + 14 + 4;
            using MemoryStream workbook = new();
            WriteRecord(workbook, 0x0809, [.. U16(0x0600), .. U16(0x0005), .. new byte[12]]);
            WriteRecord(workbook, 0x0085, BoundSheet(globalsLength, "S1"));
            WriteRecord(workbook, 0x000A, []);
            workbook.Write(sheetBytes);
            return BuildOle(workbook.ToArray());
        }

        // OLE header/layout offsets, matching Header() and BuildOle() below. Used by error-path
        // tests to corrupt one field of an otherwise-valid container.
        internal const int SectorShiftOffset = 0x1E;     // log2(sector size); valid is 9 -> 512
        internal const int FatSectorCountOffset = 0x2C;  // header DIFAT lists this many FAT sectors
        internal const int MiniCutoffOffset = 0x38;      // header's mini stream cutoff field (Int32)
        internal const int MiniFatSectorCountOffset = 0x40; // number of mini-FAT sectors (Int32)
        internal const int SignatureOffset = 0x00;
        // Directory is sector 1: header (512) + FAT sector (512) = byte 1024. The Workbook entry
        // is the second 128-byte directory entry, so its UTF-16 name starts at 1024 + 128.
        internal const int WorkbookEntryNameOffset = 1024 + 128;
        // The Workbook entry's Int64 Size field (see WriteDirectoryEntry: offset 120 within the entry).
        internal const int WorkbookSizeOffset = 1024 + 128 + 120;
        // The Root Entry's Int64 Size field — the first 128-byte directory entry, so no +128 offset.
        internal const int RootEntrySizeOffset = 1024 + 120;

        // A valid single-sheet workbook with `replacement` overwritten at `offset`.
        internal static MemoryStream BuildPatched(int offset, params byte[] replacement)
        {
            byte[] bytes = Build(sheets: [("S1", [["A"]])]).ToArray();
            replacement.CopyTo(bytes, offset);
            return new MemoryStream(bytes);
        }

        internal static byte[] LE32(int value)
        {
            return I32(value);
        }

        internal static byte[] LE16(int value)
        {
            return U16(value);
        }

        internal static byte[] LE64(long value)
        {
            byte[] bytes = new byte[8];
            BinaryPrimitives.WriteInt64LittleEndian(bytes, value);
            return bytes;
        }

        internal static byte[] RawLabel(int row, int col, string value)
        {
            return [.. U16(row), .. U16(col), .. U16(0), .. BiffString(value)];
        }

        internal static byte[] RawRowOnly(int row)
        {
            return U16(row);
        }

        // Builds a workbook whose SST is supplied pre-framed (a 0x00FC record plus any 0x003C CONTINUE
        // records), so a test can force a shared string's character array to straddle a CONTINUE
        // boundary. `labelSstCount` LabelSst cells (indices 0..count-1) are written to sheet "S1" so the
        // strings are actually resolved on read. Mirrors BuildGlobals' record order (SST before BoundSheet).
        internal static MemoryStream BuildRawSst(byte[] framedSst, int labelSstCount)
        {
            using MemoryStream sheet = new();
            WriteRecord(sheet, 0x0809, [.. U16(0x0600), .. U16(0x0010), .. new byte[12]]);
            for (int i = 0; i < labelSstCount; i++)
            {
                WriteRecord(sheet, 0x00FD, [.. U16(0), .. U16(i), .. U16(0), .. I32(i)]);
            }
            WriteRecord(sheet, 0x000A, []);
            byte[] sheetBytes = sheet.ToArray();

            int globalsLength = 82 + framedSst.Length + EncodedByteCount("S1");

            using MemoryStream globals = new();
            WriteRecord(globals, 0x0809, [.. U16(0x0600), .. U16(0x0005), .. new byte[12]]);
            WriteRecord(globals, 0x00E0, Xf(0));
            WriteRecord(globals, 0x00E0, Xf(14));
            globals.Write(framedSst);
            WriteRecord(globals, 0x0085, BoundSheet(globalsLength, "S1"));
            WriteRecord(globals, 0x000A, []);

            using MemoryStream workbook = new();
            workbook.Write(globals.ToArray());
            workbook.Write(sheetBytes);
            return BuildOle(workbook.ToArray());
        }

        // Frames a 0x00FC SST record carrying `firstRegion`, then one 0x003C CONTINUE record carrying
        // `continueRegion`. Both regions are raw string bytes; the SST's 8-byte cstTotal/cstUnique header
        // is prepended here.
        internal static byte[] FrameSstWithContinue(int cstTotal, int cstUnique, byte[] firstRegion, byte[] continueRegion)
        {
            using MemoryStream ms = new();
            WriteRecord(ms, 0x00FC, [.. I32(cstTotal), .. I32(cstUnique), .. firstRegion]);
            WriteRecord(ms, 0x003C, continueRegion);
            return ms.ToArray();
        }

        private static byte[] BuildGlobals(
            (string Name, object?[][] Rows)[] sheets,
            byte[][] sheetStreams,
            List<string> sharedStrings,
            bool date1904,
            string? customDateFormat)
        {
            int globalsLength = 4 + 16; // BOF
            byte[]? format = customDateFormat is null ? null : Format(165, customDateFormat);
            if (format is not null)
            {
                globalsLength += 4 + format.Length;
            }
            globalsLength += 4 + 20; // default XF
            globalsLength += 4 + 20; // date XF
            if (date1904)
            {
                globalsLength += 4 + 2;
            }
            byte[] sst = BuildSst(sharedStrings);
            if (sst.Length > 0)
            {
                globalsLength += 4 + sst.Length;
            }
            foreach ((string name, _) in sheets)
            {
                globalsLength += 4 + 6 + EncodedByteCount(name);
            }
            globalsLength += 4; // EOF

            using MemoryStream globals = new();
            WriteRecord(globals, 0x0809, [.. U16(0x0600), .. U16(0x0005), .. new byte[12]]);
            if (format is not null)
            {
                WriteRecord(globals, 0x041E, format);
            }
            WriteRecord(globals, 0x00E0, Xf(0));
            WriteRecord(globals, 0x00E0, Xf(format is null ? 14 : 165));
            if (date1904)
            {
                WriteRecord(globals, 0x0022, U16(1));
            }
            if (sst.Length > 0)
            {
                WriteRecord(globals, 0x00FC, sst);
            }

            int offset = globalsLength;
            for (int i = 0; i < sheets.Length; i++)
            {
                WriteRecord(globals, 0x0085, BoundSheet(offset, sheets[i].Name));
                offset += sheetStreams[i].Length;
            }
            WriteRecord(globals, 0x000A, []);
            return globals.ToArray();
        }

        private static byte[] BuildSheet(object?[][] rows, List<string> sharedStrings, Dictionary<string, int> sharedIndexes, bool date1904)
        {
            using MemoryStream sheet = new();
            WriteRecord(sheet, 0x0809, [.. U16(0x0600), .. U16(0x0010), .. new byte[12]]);
            for (int r = 0; r < rows.Length; r++)
            {
                object?[] row = rows[r];
                for (int c = 0; c < row.Length; c++)
                {
                    WriteCell(sheet, r, c, row[c], sharedStrings, sharedIndexes, date1904);
                }
            }
            WriteRecord(sheet, 0x000A, []);
            return sheet.ToArray();
        }

        private static void WriteCell(
            MemoryStream sheet,
            int row,
            int col,
            object? value,
            List<string> sharedStrings,
            Dictionary<string, int> sharedIndexes,
            bool date1904)
        {
            switch (value)
            {
                case null:
                    return;
                case XlsSharedString s:
                    if (!sharedIndexes.TryGetValue(s.Value, out int index))
                    {
                        index = sharedStrings.Count;
                        sharedStrings.Add(s.Value);
                        sharedIndexes.Add(s.Value, index);
                    }
                    WriteRecord(sheet, 0x00FD, [.. U16(row), .. U16(col), .. U16(0), .. I32(index)]);
                    break;
                case XlsSharedIndex s:
                    WriteRecord(sheet, 0x00FD, [.. U16(row), .. U16(col), .. U16(0), .. I32(s.Index)]);
                    break;
                case string s:
                    WriteRecord(sheet, 0x0204, [.. U16(row), .. U16(col), .. U16(0), .. BiffString(s)]);
                    break;
                case XlsUnicodeString s:
                    WriteRecord(sheet, 0x0204, [.. U16(row), .. U16(col), .. U16(0), .. BiffUnicodeString(s.Value)]);
                    break;
                case XlsCompressedBytes s:
                    WriteRecord(sheet, 0x0204, [.. U16(row), .. U16(col), .. U16(0), .. BiffCompressedBytes(s.Bytes)]);
                    break;
                case int i:
                    WriteNumber(sheet, row, col, 0, i);
                    break;
                case double d:
                    WriteNumber(sheet, row, col, 0, d);
                    break;
                case bool b:
                    WriteRecord(sheet, 0x0205, [.. U16(row), .. U16(col), .. U16(0), (byte)(b ? 1 : 0), 0]);
                    break;
                case XlsDate d:
                    double serial = date1904 ? d.Value.ToOADate() - 1462.0 : d.Value.ToOADate();
                    WriteNumber(sheet, row, col, 1, serial);
                    break;
                case XlsError e:
                    WriteRecord(sheet, 0x0205, [.. U16(row), .. U16(col), .. U16(0), e.Code, 1]);
                    break;
                case XlsFormula f:
                    byte[] data = [.. U16(row), .. U16(col), .. U16(0), .. Double(f.Value), .. new byte[8]];
                    WriteRecord(sheet, 0x0006, data);
                    break;
                case XlsFormulaBool f:
                    WriteRecord(sheet, 0x0006, [.. U16(row), .. U16(col), .. U16(0), .. FormulaSpecial(1, (byte)(f.Value ? 1 : 0)), .. new byte[8]]);
                    break;
                case XlsFormulaError f:
                    WriteRecord(sheet, 0x0006, [.. U16(row), .. U16(col), .. U16(0), .. FormulaSpecial(2, f.Code), .. new byte[8]]);
                    break;
                case XlsRkInt rk:
                    WriteRecord(sheet, 0x027E, [.. U16(row), .. U16(col), .. U16(0), .. U32(EncodeRkInt(rk.Value))]);
                    break;
                case XlsRkRaw rk:
                    WriteRecord(sheet, 0x027E, [.. U16(row), .. U16(col), .. U16(0), .. U32(rk.Value)]);
                    break;
                case XlsMulRk mul:
                    WriteMulRk(sheet, row, col, mul.Values);
                    break;
                case XlsBlank:
                    WriteRecord(sheet, 0x0201, [.. U16(row), .. U16(col), .. U16(0)]);
                    break;
                case XlsMulBlank mb:
                    WriteMulBlank(sheet, row, col, mb.Count);
                    break;
                case XlsAt at:
                    WriteCell(sheet, row, at.Column, at.Value, sharedStrings, sharedIndexes, date1904);
                    break;
                default:
                    throw new NotSupportedException(value.GetType().FullName);
            }
        }

        private static byte[] BuildSst(List<string> sharedStrings)
        {
            if (sharedStrings.Count == 0)
            {
                return [];
            }
            using MemoryStream sst = new();
            sst.Write(I32(sharedStrings.Count));
            sst.Write(I32(sharedStrings.Count));
            foreach (ref readonly var value in CollectionsMarshal.AsSpan(sharedStrings))
            {
                sst.Write(BiffString(value));
            }
            return sst.ToArray();
        }

        private static void WriteNumber(MemoryStream sheet, int row, int col, int xf, double value)
        {
            WriteRecord(sheet, 0x0203, [.. U16(row), .. U16(col), .. U16(xf), .. Double(value)]);
        }

        private static byte[] BoundSheet(int offset, string name)
        {
            return [.. I32(offset), 0, 0, .. BiffShortString(name)];
        }

        private static byte[] Xf(int format)
        {
            byte[] data = new byte[20];
            BinaryPrimitives.WriteUInt16LittleEndian(data.AsSpan(2), (ushort)format);
            return data;
        }

        private static byte[] Format(int id, string value)
        {
            return [.. U16(id), .. BiffString(value)];
        }

        private static byte[] BiffShortString(string value)
        {
            if (CanCompress(value))
            {
                byte[] text = EncodeCompressed(value);
                return [(byte)value.Length, 0, .. text];
            }
            return [(byte)value.Length, 1, .. Encoding.Unicode.GetBytes(value)];
        }

        private static byte[] BiffString(string value)
        {
            return CanCompress(value)
                ? [.. U16(value.Length), 0, .. EncodeCompressed(value)]
                : BiffUnicodeString(value);
        }

        private static byte[] BiffUnicodeString(string value)
        {
            return [.. U16(value.Length), 1, .. Encoding.Unicode.GetBytes(value)];
        }

        private static byte[] BiffCompressedBytes(byte[] value)
        {
            return [.. U16(value.Length), 0, .. value];
        }

        private static byte[] FormulaSpecial(byte kind, byte value)
        {
            return [kind, 0, value, 0, 0, 0, 0xFF, 0xFF];
        }

        private static uint EncodeRkInt(int value)
        {
            return ((uint)value << 2) | 0x02;
        }

        private static void WriteMulRk(MemoryStream sheet, int row, int col, int[] values)
        {
            using MemoryStream data = new();
            data.Write(U16(row));
            data.Write(U16(col));
            foreach (int v in values)
            {
                data.Write(U16(0));
                data.Write(U32(EncodeRkInt(v)));
            }
            data.Write(U16(col + values.Length - 1));
            WriteRecord(sheet, 0x00BD, data.ToArray());
        }

        private static void WriteMulBlank(MemoryStream sheet, int row, int col, int count)
        {
            using MemoryStream data = new();
            data.Write(U16(row));
            data.Write(U16(col));
            for (int i = 0; i < count; i++)
            {
                data.Write(U16(0));
            }
            data.Write(U16(col + count - 1));
            WriteRecord(sheet, 0x00BE, data.ToArray());
        }

        private static int EncodedByteCount(string value)
        {
            return 2 + (CanCompress(value) ? value.Length : value.Length * 2);
        }

        private static byte[] EncodeCompressed(string value)
        {
            byte[] bytes = new byte[value.Length];
            for (int i = 0; i < value.Length; i++)
            {
                bytes[i] = value[i] <= 0xFF ? (byte)value[i] : (byte)'?';
            }
            return bytes;
        }

        private static bool CanCompress(string value)
        {
            for (int i = 0; i < value.Length; i++)
            {
                if (value[i] > 0xFF)
                {
                    return false;
                }
            }
            return true;
        }

        private static MemoryStream BuildOle(byte[] workbook)
        {
            int workbookSize = Math.Max(RoundUp(workbook.Length, SectorSize), 4096 + SectorSize);
            int workbookSectorCount = workbookSize / SectorSize;
            int totalSectors = 2 + workbookSectorCount;
            if (totalSectors > 128)
            {
                throw new NotSupportedException("Test workbook too large for the simple OLE builder.");
            }

            byte[] fat = [.. Enumerable.Repeat(FreeSector, 128).SelectMany(I32)];
            WriteI32(fat, 0, FatSector);
            WriteI32(fat, 4, EndOfChain);
            for (int i = 0; i < workbookSectorCount; i++)
            {
                int sector = 2 + i;
                int next = i == workbookSectorCount - 1 ? EndOfChain : sector + 1;
                WriteI32(fat, sector * 4, next);
            }

            MemoryStream ole = new();
            ole.Write(Header());
            ole.Write(fat);
            ole.Write(Directory(workbookSize));
            ole.Write(workbook);
            ole.Write(new byte[workbookSize - workbook.Length]);
            ole.Position = 0;
            return ole;
        }

        private static byte[] Header()
        {
            byte[] header = new byte[SectorSize];
            byte[] sig = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
            sig.CopyTo(header, 0);
            WriteU16(header, 0x1A, 0x003E);
            WriteU16(header, 0x1C, 0x0003);
            WriteU16(header, 0x1E, 9);
            WriteU16(header, 0x20, 6);
            WriteI32(header, 0x2C, 1);
            WriteI32(header, 0x30, 1);
            WriteI32(header, 0x38, 4096);
            WriteI32(header, 0x3C, EndOfChain);
            WriteI32(header, 0x40, 0);
            WriteI32(header, 0x44, EndOfChain);
            WriteI32(header, 0x48, 0);
            for (int i = 0x4C; i < SectorSize; i += 4)
            {
                WriteI32(header, i, FreeSector);
            }
            WriteI32(header, 0x4C, 0);
            return header;
        }

        private static byte[] Directory(int workbookSize)
        {
            byte[] dir = new byte[SectorSize];
            WriteDirectoryEntry(dir.AsSpan(0, 128), "Root Entry", 5, EndOfChain, 0, child: 1);
            WriteDirectoryEntry(dir.AsSpan(128, 128), "Workbook", 2, 2, workbookSize, child: EndOfChain);
            return dir;
        }

        private static void WriteDirectoryEntry(Span<byte> entry, string name, byte type, int startSector, long size, int child)
        {
            Encoding.Unicode.GetBytes(name + '\0').CopyTo(entry);
            WriteU16(entry, 64, (ushort)((name.Length + 1) * 2));
            entry[66] = type;
            entry[67] = 1;
            WriteI32(entry, 68, EndOfChain);
            WriteI32(entry, 72, EndOfChain);
            WriteI32(entry, 76, child);
            WriteI32(entry, 116, startSector);
            BinaryPrimitives.WriteInt64LittleEndian(entry[120..], size);
        }

        private static void WriteRecord(MemoryStream stream, int id, byte[] data)
        {
            stream.Write(U16(id));
            stream.Write(U16(data.Length));
            stream.Write(data);
        }

        private static byte[] Record(int id, byte[] data)
        {
            using MemoryStream ms = new();
            WriteRecord(ms, id, data);
            return ms.ToArray();
        }

        private static int RoundUp(int value, int multiple)
        {
            return (value + multiple - 1) / multiple * multiple;
        }

        private static byte[] U16(int value)
        {
            byte[] bytes = new byte[2];
            WriteU16(bytes, 0, (ushort)value);
            return bytes;
        }

        private static byte[] I32(int value)
        {
            byte[] bytes = new byte[4];
            WriteI32(bytes, 0, value);
            return bytes;
        }

        private static byte[] U32(uint value)
        {
            byte[] bytes = new byte[4];
            BinaryPrimitives.WriteUInt32LittleEndian(bytes, value);
            return bytes;
        }

        private static byte[] Double(double value)
        {
            byte[] bytes = new byte[8];
            BinaryPrimitives.WriteDoubleLittleEndian(bytes, value);
            return bytes;
        }

        private static void WriteU16(Span<byte> bytes, int offset, ushort value)
        {
            BinaryPrimitives.WriteUInt16LittleEndian(bytes.Slice(offset, 2), value);
        }

        private static void WriteI32(Span<byte> bytes, int offset, int value)
        {
            BinaryPrimitives.WriteInt32LittleEndian(bytes.Slice(offset, 4), value);
        }
    }
}
