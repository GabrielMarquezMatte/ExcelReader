using System.Runtime.InteropServices;
using ExcelReader.NetFramework.Interop;

namespace ExcelReader.NetFramework.Writer.Internal
{
    internal sealed unsafe class NativeWriteTable : IDisposable
    {
        private readonly List<IntPtr> _blocks = new List<IntPtr>();

        internal NativeWriteTable()
        {
        }

        internal XlTable* Table { get; private set; }

        internal XlColumnSpec* Specs { get; private set; }

        internal XlCellStyle** ColumnStyles { get; private set; }

        internal XlCellStyle* HeaderStyle { get; private set; }

        internal static NativeWriteTable Create<T>(WriteMap<T> map, T[] rows, CellStyle? headerStyle = null)
        {
            NativeWriteTable table = new NativeWriteTable();
            try
            {
                table.Build(map, rows, headerStyle);
                return table;
            }
            catch
            {
                table.Dispose();
                throw;
            }
        }

        internal byte* Allocate(long bytes)
        {
            _blocks.Add(IntPtr.Zero);
            IntPtr block = Marshal.AllocHGlobal(new IntPtr(Math.Max(bytes, 1)));
            _blocks[_blocks.Count - 1] = block;
            return (byte*)block;
        }

        internal XlCellStyle* Style(CellStyle style)
        {
            XlCellStyle* native = (XlCellStyle*)Allocate(sizeof(XlCellStyle));
            *native = new XlCellStyle
            {
                StructSize = sizeof(XlCellStyle),
                FontFlags = (style.Bold ? Xl.XL_FONT_BOLD : 0) | (style.Italic ? Xl.XL_FONT_ITALIC : 0) | (style.Underline ? Xl.XL_FONT_UNDERLINE : 0) | (style.Strikethrough ? Xl.XL_FONT_STRIKETHROUGH : 0),
                FontSize = style.FontSize ?? 0,
                FontColor = style.FontColor,
                FillColor = style.FillColor,
                HorizontalAlignment = (int)style.HorizontalAlignment,
                VerticalAlignment = (int)style.VerticalAlignment,
                WrapText = style.WrapText ? 1 : 0,
                Indent = style.Indent,
                Rotation = style.Rotation,
                Left = Border(style.Left),
                Right = Border(style.Right),
                Top = Border(style.Top),
                Bottom = Border(style.Bottom),
            };
            native->NumberFormat = Text(style.NumberFormat, out native->NumberFormatLen);
            native->FontName = Text(style.FontName, out native->FontNameLen);
            return native;
        }

        internal byte* AllocateValidity(int rows)
        {
            int length = (rows + 7) / 8;
            byte* validity = Allocate(length);
            for (int i = 0; i < length; i++)
            {
                validity[i] = 0;
            }

            return validity;
        }

        public void Dispose()
        {
            foreach (IntPtr block in _blocks)
            {
                if (block != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(block);
                }
            }

            _blocks.Clear();
            Table = null;
            Specs = null;
            ColumnStyles = null;
            HeaderStyle = null;
        }

        private static XlBorder Border(CellBorder? border)
        {
            if (border is null)
            {
                return new XlBorder { Style = 0, Color = -1 };
            }

            return new XlBorder { Style = (int)border.Style, Color = border.Color };
        }

        private byte* Text(string? value, out int length)
        {
            length = 0;
            if (value is null)
            {
                return null;
            }

            byte[] bytes = Utf8.NonNullBytes(value, out length);
            byte* text = Allocate(bytes.Length);
            Marshal.Copy(bytes, 0, (IntPtr)text, bytes.Length);
            return text;
        }

        private void Build<T>(WriteMap<T> map, T[] rows, CellStyle? headerStyle)
        {
            ColumnWriter<T>[] writers = map.Columns;
            int count = writers.Length;
            XlColumnSpec* specs = (XlColumnSpec*)Allocate((long)count * sizeof(XlColumnSpec));
            XlColumn* columns = (XlColumn*)Allocate((long)count * sizeof(XlColumn));
            for (int i = 0; i < count; i++)
            {
                specs[i] = Spec(writers[i], i);
                writers[i].Fill(rows, columns + i, this);
            }

            if (writers.Any(w => w.Style is not null))
            {
                XlCellStyle** styles = (XlCellStyle**)Allocate((long)count * sizeof(XlCellStyle*));
                for (int i = 0; i < count; i++)
                {
                    CellStyle? style = writers[i].Style;
                    styles[i] = style is null ? null : Style(style);
                }

                ColumnStyles = styles;
            }

            if (headerStyle is not null)
            {
                HeaderStyle = Style(headerStyle);
            }

            XlTable* table = (XlTable*)Allocate(sizeof(XlTable));
            *table = new XlTable { ColumnCount = count, RowCount = rows.Length, Columns = columns };
            Specs = specs;
            Table = table;
        }

        private XlColumnSpec Spec<T>(ColumnWriter<T> writer, int index)
        {
            byte[] name = Utf8.Bytes(writer.Header);
            byte** names = (byte**)Allocate(sizeof(byte*));
            int* nameLengths = (int*)Allocate(sizeof(int));
            byte* text = Allocate(name.Length);
            Marshal.Copy(name, 0, (IntPtr)text, name.Length);
            *names = text;
            *nameLengths = name.Length;
            return new XlColumnSpec { Names = names, NameLens = nameLengths, NameCount = 1, Index = index, Type = writer.NativeType, Nullable = writer.Nullable ? 1 : 0 };
        }
    }
}
