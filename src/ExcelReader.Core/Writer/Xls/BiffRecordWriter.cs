using System.Buffers;
using System.Buffers.Binary;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Xls
{
    internal static class BiffRecordWriter
    {
        internal static void WriteBof(BiffBuffer buffer, int substreamType)
        {
            int len = buffer.BeginRecord(BiffRecord.Bof);
            buffer.WriteU16(BiffRecord.Biff8Version);
            buffer.WriteU16(substreamType);
            // [MS-XLS] requires rupYear to be 1996 or 1997, and verLowestBiff to be 6.
            buffer.WriteU16(0x0DBB);
            buffer.WriteU16(0x07CC);
            buffer.WriteI32(0);
            buffer.WriteI32(6);
            buffer.EndRecord(len);
        }

        internal static void WriteEof(BiffBuffer buffer)
        {
            int len = buffer.BeginRecord(BiffRecord.Eof);
            buffer.EndRecord(len);
        }

        internal static void WriteCodePage(BiffBuffer buffer, int codePage)
        {
            int len = buffer.BeginRecord(BiffRecord.CodePage);
            buffer.WriteU16(codePage);
            buffer.EndRecord(len);
        }

        internal static void WriteInterfaceHdr(BiffBuffer buffer, int codePage)
        {
            int len = buffer.BeginRecord(BiffRecord.InterfaceHdr);
            buffer.WriteU16(codePage);
            buffer.EndRecord(len);
        }

        internal static void WriteMms(BiffBuffer buffer)
        {
            int len = buffer.BeginRecord(BiffRecord.Mms);
            buffer.WriteU16(0);
            buffer.EndRecord(len);
        }

        internal static void WriteInterfaceEnd(BiffBuffer buffer)
        {
            int len = buffer.BeginRecord(BiffRecord.InterfaceEnd);
            buffer.EndRecord(len);
        }

        internal static void WriteWriteAccess(BiffBuffer buffer)
        {
            int len = buffer.BeginRecord(BiffRecord.WriteAccess);
            ReadOnlySpan<byte> name = "ExcelReader"u8;
            buffer.WriteU16(name.Length);
            buffer.WriteByte(0);
            buffer.Write(name);
            for (int i = 3 + name.Length; i < 112; i++)
            {
                buffer.WriteByte((byte)' ');
            }
            buffer.EndRecord(len);
        }

        internal static void WriteDsf(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.Dsf, 0);
        }

        internal static void WriteTabId(BiffBuffer buffer, int sheetCount)
        {
            int len = buffer.BeginRecord(BiffRecord.TabId);
            for (int i = 1; i <= sheetCount; i++)
            {
                buffer.WriteU16(i);
            }
            buffer.EndRecord(len);
        }

        internal static void WriteFnGroupCount(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.FnGroupCount, 17);
        }
        internal static void WriteWindow1(BiffBuffer buffer, int activeSheet)
        {
            int len = buffer.BeginRecord(BiffRecord.Window1);
            buffer.WriteU16(0x7FFF);
            buffer.WriteU16(0x7FFF);
            buffer.WriteU16(0x7080);
            buffer.WriteU16(0x3930);
            buffer.WriteU16(0x0038);
            buffer.WriteU16(activeSheet);
            buffer.WriteU16(0);
            buffer.WriteU16(1);
            buffer.WriteU16(600);
            buffer.EndRecord(len);
        }

        internal static void WriteU16Records(BiffBuffer buffer, ReadOnlySpan<int> ids, int value)
        {
            foreach (int id in ids)
            {
                WriteU16Record(buffer, id, value);
            }
        }

        internal static void WriteBackup(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.Backup, 0);
        }

        internal static void WriteHideObj(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.HideObj, 0);
        }

        internal static void WriteDate1904(BiffBuffer buffer, bool date1904)
        {
            int len = buffer.BeginRecord(BiffRecord.Date1904);
            buffer.WriteU16(date1904 ? 1 : 0);
            buffer.EndRecord(len);
        }

        internal static void WritePrecision(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.Precision, 1);
        }

        internal static void WriteRefreshAll(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.RefreshAll, 0);
        }

        internal static void WriteBookBool(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.BookBool, 0);
        }

        internal static void WriteFont(BiffBuffer buffer, in FontKey font, IReadOnlyList<ExcelColor> palette)
        {
            int len = buffer.BeginRecord(BiffRecord.Font);
            buffer.WriteU16(font.Twips);
            buffer.WriteU16(font.Flags);
            buffer.WriteU16(ColorIndex(palette, font.Color, AutomaticFontColor));
            buffer.WriteU16(font.Bold ? 700 : 400);
            buffer.WriteU16(0);
            buffer.WriteByte(font.Underline ? (byte)1 : (byte)0);
            buffer.WriteByte(0);
            buffer.WriteByte(0);
            buffer.WriteByte(0);
            BiffStringEncoder.WriteShort(buffer, font.Name ?? "Arial");
            buffer.EndRecord(len);
        }

        private static ReadOnlySpan<byte> FirstPayload => [0xF5, 0xFF, 0x20, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x20];

        internal static void WriteStyleXf(BiffBuffer buffer, int formatIndex)
        {
            int len = buffer.BeginRecord(BiffRecord.Xf);
            buffer.WriteU16(0);
            buffer.WriteU16(formatIndex);
            buffer.Write(FirstPayload);
            buffer.EndRecord(len);
        }

        internal static void WriteCellXf(BiffBuffer buffer, int fontIndex, int formatIndex, in CellStyle style, IReadOnlyList<ExcelColor> palette)
        {
            int lineStyles = LineStyle(style.Left) | (LineStyle(style.Right) << 4) | (LineStyle(style.Top) << 8) | (LineStyle(style.Bottom) << 12);
            int used = 0;
            used |= fontIndex != 0 ? XfUsedFont : 0;
            used |= StyleAlignment.HasAlignment(style) ? XfUsedAlignment : 0;
            used |= lineStyles != 0 ? XfUsedBorder : 0;
            used |= style.FillColor is not null ? XfUsedPattern : 0;
            int len = buffer.BeginRecord(BiffRecord.Xf);
            buffer.WriteU16(fontIndex);
            buffer.WriteU16(formatIndex);
            buffer.WriteU16(XfLocked);
            buffer.WriteByte((byte)Alignment(style));
            buffer.WriteByte((byte)StyleAlignment.TextRotation(style.Rotation));
            buffer.WriteByte((byte)style.Indent);
            buffer.WriteByte((byte)used);
            buffer.WriteI32(lineStyles | (BorderColor(palette, style.Left) << 16) | (BorderColor(palette, style.Right) << 23));
            buffer.WriteI32(BorderColor(palette, style.Top) | (BorderColor(palette, style.Bottom) << 7) | XfHasXfExt | FillPattern(style.FillColor));
            buffer.WriteU16(FillColors(palette, style.FillColor));
            buffer.EndRecord(len);
        }

        internal static void WritePalette(BiffBuffer buffer, IReadOnlyList<ExcelColor> palette)
        {
            int len = buffer.BeginRecord(BiffRecord.Palette);
            buffer.WriteU16(DefaultPalette.Length);
            for (int i = 0; i < DefaultPalette.Length; i++)
            {
                ExcelColor color = i < palette.Count ? palette[i] : ExcelColor.FromRgb(DefaultPalette[i]);
                buffer.WriteByte(color.R);
                buffer.WriteByte(color.G);
                buffer.WriteByte(color.B);
                buffer.WriteByte(0);
            }
            buffer.EndRecord(len);
        }

        private static int LineStyle(CellBorder? side)
        {
            return (int)(side?.Style ?? ExcelBorderStyle.None);
        }

        private static int BorderColor(IReadOnlyList<ExcelColor> palette, CellBorder? side)
        {
            if (side is not CellBorder border || border.Style == ExcelBorderStyle.None)
            {
                return 0;
            }
            return ColorIndex(palette, border.Color, AutomaticColor);
        }

        private static int FillPattern(ExcelColor? fill)
        {
            return fill is null ? 0 : XfSolidFill;
        }

        private static int FillColors(IReadOnlyList<ExcelColor> palette, ExcelColor? fill)
        {
            if (fill is null)
            {
                return AutomaticColor | ((AutomaticColor + 1) << 7);
            }
            return ColorIndex(palette, fill, AutomaticColor) | (AutomaticColor << 7);
        }

        private static int Alignment(in CellStyle style)
        {
            int alignment = (int)style.HorizontalAlignment | (StyleAlignment.BiffVerticalAlignment(style.VerticalAlignment) << 4);
            alignment |= style.WrapText ? XfWrap : 0;
            return alignment;
        }

        // Palette color i is icv 8 + i; StyleTable puts every style color in the palette.
        private static int ColorIndex(IReadOnlyList<ExcelColor> palette, ExcelColor? color, int automatic)
        {
            if (color is not ExcelColor rgb)
            {
                return automatic;
            }
            int i = 0;
            while (palette[i] != rgb)
            {
                i++;
            }
            return 8 + i;
        }

        private const int AutomaticFontColor = 0x7FFF;
        private const int AutomaticColor = 64;
        private const int XfLocked = 0x0001;
        private const int XfWrap = 0x08;
        private const int XfUsedFont = 0x08;
        private const int XfUsedAlignment = 0x10;
        private const int XfUsedBorder = 0x20;
        private const int XfUsedPattern = 0x40;
        private const int XfSolidFill = 1 << 26;
        // Every cell XF Excel writes sets fHasXFExt; the default cell XF here always has.
        private const int XfHasXfExt = 1 << 25;

        // Excel's default BIFF8 palette, icv 8 to 63.
        private static ReadOnlySpan<int> DefaultPalette =>
        [
            0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
            0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
            0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
            0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
            0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
            0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
            0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333,
        ];

        internal static void WriteFormat(BiffBuffer buffer, int formatIndex, ReadOnlySpan<char> formatCode)
        {
            int len = buffer.BeginRecord(BiffRecord.Format);
            buffer.WriteU16(formatIndex);
            BiffStringEncoder.WriteLong(buffer, formatCode);
            buffer.EndRecord(len);
        }

        internal static void WriteStyle(BiffBuffer buffer)
        {
            int len = buffer.BeginRecord(BiffRecord.Style);
            buffer.WriteU16(0x8000);
            buffer.WriteByte(0);
            buffer.WriteByte(0xFF);
            buffer.EndRecord(len);
        }

        internal static void WriteBoundSheet(BiffBuffer buffer, int sheetOffset, ReadOnlySpan<char> name,
            ExcelSheetVisibility visibility = ExcelSheetVisibility.Visible)
        {
            int len = buffer.BeginRecord(BiffRecord.BoundSheet);
            buffer.WriteI32(sheetOffset);
            buffer.WriteU16((ushort)visibility);
            BiffStringEncoder.WriteShort(buffer, name);
            buffer.EndRecord(len);
        }

        // ponytail: field layout (colFirst/colLast/coldx/ixfe/grbit/reserved, all u16) follows the
        internal static void WriteColInfo(BiffBuffer buffer, int columnIndex, int width256ths, int xf)
        {
            int len = buffer.BeginRecord(BiffRecord.ColInfo);
            buffer.WriteU16(columnIndex);
            buffer.WriteU16(columnIndex);
            buffer.WriteU16(width256ths);
            buffer.WriteU16(xf);
            buffer.WriteU16(0);
            buffer.WriteU16(0);
            buffer.EndRecord(len);
        }

        internal static void WriteDimension(BiffBuffer buffer, int rowCount, int colCount)
        {
            int len = buffer.BeginRecord(BiffRecord.Dimension);
            buffer.WriteI32(0);
            buffer.WriteI32(rowCount);
            buffer.WriteU16(0);
            buffer.WriteU16(colCount);
            buffer.WriteU16(0);
            buffer.EndRecord(len);
        }
        internal static void WriteWindow2(BiffBuffer buffer, bool selected)
        {
            int len = buffer.BeginRecord(BiffRecord.Window2);
            buffer.WriteU16(selected ? 0x06B6 : 0x00B6);
            buffer.WriteU16(0);
            buffer.WriteU16(0);
            buffer.WriteU16(0x0040);
            buffer.WriteU16(0);
            buffer.WriteU16(0);
            buffer.WriteU16(0);
            buffer.WriteI32(0);
            buffer.EndRecord(len);
        }

        internal static void WriteSheetSettings(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.CalcMode, 1);
            WriteU16Record(buffer, BiffRecord.CalcCount, 100);
            WriteU16Record(buffer, BiffRecord.CalcRefMode, 1);
            WriteU16Record(buffer, BiffRecord.CalcIter, 0);
            int delta = buffer.BeginRecord(BiffRecord.CalcDelta);
            buffer.WriteDouble(0.001);
            buffer.EndRecord(delta);
            WriteU16Record(buffer, BiffRecord.CalcSaveRecalc, 1);
            WriteU16Record(buffer, BiffRecord.PrintRowCol, 0);
            WriteU16Record(buffer, BiffRecord.PrintGrid, 0);
            WriteU16Record(buffer, BiffRecord.GridSet, 1);
            int guts = buffer.BeginRecord(BiffRecord.Guts);
            buffer.WriteI32(0);
            buffer.WriteI32(0);
            buffer.EndRecord(guts);
            int rowHeight = buffer.BeginRecord(BiffRecord.DefaultRowHeight);
            buffer.WriteU16(0);
            buffer.WriteU16(DefaultRowHeightTwips);
            buffer.EndRecord(rowHeight);
            WriteU16Record(buffer, BiffRecord.WsBool, 0x04C1);
        }

        internal const int DefaultRowHeightTwips = 255;

        internal static void WriteDefColWidth(BiffBuffer buffer)
        {
            WriteU16Record(buffer, BiffRecord.DefColWidth, 8);
        }

        internal static void WriteRow(BiffBuffer buffer, int row, int colMic, int colMac)
        {
            int len = buffer.BeginRecord(BiffRecord.Row);
            buffer.WriteU16(row);
            buffer.WriteU16(colMic);
            buffer.WriteU16(colMac);
            buffer.WriteU16(DefaultRowHeightTwips);
            buffer.WriteI32(0);
            buffer.WriteU16(0x0100);
            buffer.WriteU16(GeneralStyleXf);
            buffer.EndRecord(len);
        }

        private const int GeneralStyleXf = 15;

        // Cell records are written whole into one span: they are most of a sheet's bytes, and going
        // through BeginRecord/WriteU16/EndRecord field by field measured as most of their cost.
        internal static void WriteBlank(BiffBuffer buffer, int row, int col, int xf)
        {
            Span<byte> record = CellRecord(buffer, BiffRecord.Blank, 6, row, col, xf);
            buffer.Advance(record.Length);
        }

        internal static void WriteNumber(BiffBuffer buffer, int row, int col, int xf, double value)
        {
            CellValueGuards.ThrowIfNonFinite(value, nameof(value));
            Span<byte> record = CellRecord(buffer, BiffRecord.Number, 14, row, col, xf);
            BinaryPrimitives.WriteDoubleLittleEndian(record[10..], value);
            buffer.Advance(record.Length);
        }

        /// <summary>
        /// Opens a run of RK cells: header, row and first column now, then 6 bytes (ixfe, RK) per cell.
        /// A run of one is an RK record, whose layout is exactly that; a longer run is a MULRK record,
        /// which only adds the last column. <see cref="EndRkRun"/> fixes the id and length either way.
        /// </summary>
        internal static int BeginRkRun(BiffBuffer buffer, int row, int firstCol)
        {
            int start = buffer.Length;
            Span<byte> header = buffer.GetSpan(8);
            BinaryPrimitives.WriteUInt16LittleEndian(header[4..], (ushort)row);
            BinaryPrimitives.WriteUInt16LittleEndian(header[6..], (ushort)firstCol);
            buffer.Advance(8);
            return start;
        }

        internal static void AppendRk(BiffBuffer buffer, int xf, int rk)
        {
            Span<byte> cell = buffer.GetSpan(6);
            BinaryPrimitives.WriteUInt16LittleEndian(cell, (ushort)xf);
            BinaryPrimitives.WriteInt32LittleEndian(cell[2..], rk);
            buffer.Advance(6);
        }

        internal static void EndRkRun(BiffBuffer buffer, int start, int firstCol, int count)
        {
            Span<byte> header = buffer.Slice(start, 4);
            if (count == 1)
            {
                BinaryPrimitives.WriteUInt16LittleEndian(header, BiffRecord.Rk);
                BinaryPrimitives.WriteUInt16LittleEndian(header[2..], 10);
                return;
            }
            BinaryPrimitives.WriteUInt16LittleEndian(header, BiffRecord.MulRk);
            BinaryPrimitives.WriteUInt16LittleEndian(header[2..], (ushort)(4 + (6 * count) + 2));
            Span<byte> last = buffer.GetSpan(2);
            BinaryPrimitives.WriteUInt16LittleEndian(last, (ushort)(firstCol + count - 1));
            buffer.Advance(2);
        }

        internal static void WriteLabel(BiffBuffer buffer, int row, int col, int xf, ReadOnlySpan<char> value)
        {
            Span<byte> dest = buffer.GetSpan(13 + (value.Length * 2));
            Span<byte> chars = dest[13..];
            int charBytes;
            byte flags = 0;
            if (Ascii.FromUtf16(value, chars, out _) == OperationStatus.Done)
            {
                charBytes = value.Length;
            }
            else if (BiffStringEncoder.CanCompress(value))
            {
                charBytes = Encoding.Latin1.GetBytes(value, chars);
            }
            else
            {
                MemoryMarshal.AsBytes(value).CopyTo(chars);
                charBytes = value.Length * 2;
                flags = 1;
            }
            Span<byte> record = CellRecord(buffer, BiffRecord.Label, 9 + charBytes, row, col, xf);
            BinaryPrimitives.WriteUInt16LittleEndian(record[10..], (ushort)value.Length);
            record[12] = flags;
            buffer.Advance(record.Length);
        }

        internal static void WriteLabelSst(BiffBuffer buffer, int row, int col, int xf, int sharedStringIndex)
        {
            Span<byte> record = CellRecord(buffer, BiffRecord.LabelSst, 10, row, col, xf);
            BinaryPrimitives.WriteInt32LittleEndian(record[10..], sharedStringIndex);
            buffer.Advance(record.Length);
        }

        internal static void WriteBool(BiffBuffer buffer, int row, int col, int xf, bool value)
        {
            Span<byte> record = CellRecord(buffer, BiffRecord.BoolErr, 8, row, col, xf);
            record[10] = (byte)(value ? 1 : 0);
            record[11] = 0;
            buffer.Advance(record.Length);
        }

        /// <summary>Reserves a cell record and fills its header and row/col/xf; the caller writes the rest of the payload, then advances by the returned length.</summary>
        private static Span<byte> CellRecord(BiffBuffer buffer, int id, int payload, int row, int col, int xf)
        {
            Span<byte> record = buffer.GetSpan(4 + payload)[..(4 + payload)];
            BinaryPrimitives.WriteUInt16LittleEndian(record, (ushort)id);
            BinaryPrimitives.WriteUInt16LittleEndian(record[2..], (ushort)payload);
            BinaryPrimitives.WriteUInt16LittleEndian(record[4..], (ushort)row);
            BinaryPrimitives.WriteUInt16LittleEndian(record[6..], (ushort)col);
            BinaryPrimitives.WriteUInt16LittleEndian(record[8..], (ushort)xf);
            return record;
        }


        private static void WriteU16Record(BiffBuffer buffer, int id, int value)
        {
            int len = buffer.BeginRecord(id);
            buffer.WriteU16(value);
            buffer.EndRecord(len);
        }
    }
}
