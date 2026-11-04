using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Internal;
using ExcelReader.Native.Typed;

namespace ExcelReader.Native.Writer
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeBorderRaw
    {
        public int Style;
        public int Color;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal unsafe struct NativeCellStyleRaw
    {
        public int StructSize;
        public int NumberFormatLen;
        public byte* NumberFormat;
        public byte* FontName;
        public int FontNameLen;
        public int FontFlags;
        public double FontSize;
        public int FontColor;
        public int FillColor;
        public int HorizontalAlignment;
        public int VerticalAlignment;
        public int WrapText;
        public int Indent;
        public int Rotation;
        public NativeBorderRaw Left;
        public NativeBorderRaw Right;
        public NativeBorderRaw Top;
        public NativeBorderRaw Bottom;
    }

    internal static class NativeCellStyle
    {
        internal const int Bold = 1;
        internal const int Italic = 2;
        internal const int Underline = 4;
        internal const int Strikethrough = 8;
        private const int KnownFontFlags = Bold | Italic | Underline | Strikethrough;
        private const int UnsetColor = -1;

        internal static unsafe bool TryDecode(NativeCellStyleRaw* raw, out CellStyle? style, [NotNullWhen(false)] out string? error)
        {
            style = null;
            error = null;
            if (raw is null)
            {
                return true;
            }
            if (raw->StructSize != sizeof(NativeCellStyleRaw))
            {
                error = $"xl_cell_style.struct_size is {raw->StructSize}, but this library expects {sizeof(NativeCellStyleRaw)}.";
                return false;
            }
            int flags = raw->FontFlags;
            if ((flags & ~KnownFontFlags) != 0)
            {
                error = $"xl_cell_style.font_flags has unknown bits; got {flags}.";
                return false;
            }
            if (raw->WrapText is not (0 or 1))
            {
                error = $"xl_cell_style.wrap_text must be 0 or 1; got {raw->WrapText}.";
                return false;
            }
            if (!TryReadText(raw->NumberFormat, raw->NumberFormatLen, "number_format", out string? numberFormat, out error)
                || !TryReadText(raw->FontName, raw->FontNameLen, "font_name", out string? fontName, out error)
                || !TryReadColor(raw->FontColor, "font_color", out ExcelColor? fontColor, out error)
                || !TryReadColor(raw->FillColor, "fill_color", out ExcelColor? fillColor, out error)
                || !TryReadBorder(raw->Left, "left", out CellBorder? left, out error)
                || !TryReadBorder(raw->Right, "right", out CellBorder? right, out error)
                || !TryReadBorder(raw->Top, "top", out CellBorder? top, out error)
                || !TryReadBorder(raw->Bottom, "bottom", out CellBorder? bottom, out error))
            {
                return false;
            }
            double? fontSize = raw->FontSize == 0 ? null : raw->FontSize;
            CellStyle decoded = new()
            {
                NumberFormat = numberFormat,
                Bold = (flags & Bold) != 0,
                Italic = (flags & Italic) != 0,
                Underline = (flags & Underline) != 0,
                Strikethrough = (flags & Strikethrough) != 0,
                FontName = fontName,
                FontSize = fontSize,
                FontColor = fontColor,
                FillColor = fillColor,
                Left = left,
                Right = right,
                Top = top,
                Bottom = bottom,
                HorizontalAlignment = (ExcelHorizontalAlignment)raw->HorizontalAlignment,
                VerticalAlignment = (ExcelVerticalAlignment)raw->VerticalAlignment,
                WrapText = raw->WrapText == 1,
                Indent = raw->Indent,
                Rotation = raw->Rotation,
            };
            return TryValidate(decoded, out style, out error);
        }

        // CSV's AddStyle never validates, so validating here rejects the same values on every format.
        private static bool TryValidate(CellStyle decoded, out CellStyle? style, [NotNullWhen(false)] out string? error)
        {
            style = null;
            error = null;
            try
            {
                CellStyleValidation.Validate(decoded, "xl_cell_style");
            }
            catch (ArgumentException exception)
            {
                error = exception.Message;
                return false;
            }
            style = decoded;
            return true;
        }

        private static unsafe bool TryReadText(byte* text, int length, string field, out string? value, [NotNullWhen(false)] out string? error)
        {
            value = null;
            error = null;
            if (text is null || length == 0)
            {
                return true;
            }
            if (!TypedApi.IsValidNameLength(length))
            {
                error = $"xl_cell_style.{field}_len is out of range; got {length}.";
                return false;
            }
            value = Encoding.UTF8.GetString(text, length);
            return true;
        }

        private static bool TryReadColor(int rgb, string field, out ExcelColor? color, [NotNullWhen(false)] out string? error)
        {
            color = null;
            error = null;
            if (rgb == UnsetColor)
            {
                return true;
            }
            if (rgb is < 0 or > 0xFFFFFF)
            {
                error = $"xl_cell_style.{field} must be 0xRRGGBB or -1; got {rgb}.";
                return false;
            }
            color = ExcelColor.FromRgb(rgb);
            return true;
        }

        private static bool TryReadBorder(NativeBorderRaw raw, string field, out CellBorder? border, [NotNullWhen(false)] out string? error)
        {
            border = null;
            if (!TryReadColor(raw.Color, field + ".color", out ExcelColor? color, out error))
            {
                return false;
            }
            if (raw.Style != (int)ExcelBorderStyle.None)
            {
                border = new CellBorder((ExcelBorderStyle)raw.Style, color);
            }
            return true;
        }
    }
}
