namespace ExcelReader.Core.Writer.Internal
{
    internal static class CellStyleValidation
    {
        // BIFF12 format strings cap at 255 characters; font names at 31 (LOGFONT face-name limit).
        private const int MaxNumberFormatLength = 255;
        private const int MaxFontNameLength = 31;

        internal static void Validate(in CellStyle style, string paramName)
        {
            if (style.NumberFormat is { } format)
            {
                if (format.Length > MaxNumberFormatLength)
                {
                    throw new ArgumentException(
                        $"A number format cannot be longer than {MaxNumberFormatLength} characters ({format.Length} chars).", paramName);
                }
                if (WriterStateGuard.HasUnwritableCharacter(format))
                {
                    throw new ArgumentException("A number format cannot contain control characters, U+FFFE or U+FFFF.", paramName);
                }
            }
            if (style.FontName is { } fontName && (fontName.Length is 0 or > MaxFontNameLength || WriterStateGuard.HasUnwritableCharacter(fontName)))
            {
                throw new ArgumentException($"FontName must be 1 to {MaxFontNameLength} characters with no control characters.", paramName);
            }
            if (style.FontSize is { } size && !(size >= 1 && size <= 409))
            {
                throw new ArgumentException("FontSize must be between 1 and 409.", paramName);
            }
            if (style.HorizontalAlignment is < ExcelHorizontalAlignment.General or > ExcelHorizontalAlignment.Distributed)
            {
                throw new ArgumentException("HorizontalAlignment is not a defined value.", paramName);
            }
            if (style.VerticalAlignment is < ExcelVerticalAlignment.Bottom or > ExcelVerticalAlignment.Distributed)
            {
                throw new ArgumentException("VerticalAlignment is not a defined value.", paramName);
            }
            if (!IsDefined(style.Left) || !IsDefined(style.Right) || !IsDefined(style.Top) || !IsDefined(style.Bottom))
            {
                throw new ArgumentException("A border Style is not a defined value.", paramName);
            }
            if (style.Indent is < 0 or > 15)
            {
                throw new ArgumentException("Indent must be between 0 and 15.", paramName);
            }
            if (style.Rotation is (< -90 or > 90) && style.Rotation != CellStyle.StackedTextRotation)
            {
                throw new ArgumentException($"Rotation must be between -90 and 90, or {CellStyle.StackedTextRotation} for stacked text.", paramName);
            }
        }

        private static bool IsDefined(CellBorder? border)
            => border is null || border.Value.Style is >= ExcelBorderStyle.None and <= ExcelBorderStyle.SlantDashDot;
    }
}
