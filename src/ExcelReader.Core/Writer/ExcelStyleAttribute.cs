using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// Styles a property's data cells when <see cref="ExcelRecordLayout"/> writes it; the header cell is left alone.
    /// The style replaces any column or row style on those cells. Colors are <c>0xRRGGBB</c>, -1 for unset.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ExcelStyleAttribute : Attribute
    {
        /// <summary>Gets or sets the number-format code, or <see langword="null"/> for the general format.</summary>
        public string? NumberFormat { get; set; }

        /// <summary>Gets or sets a value indicating whether the text is bold.</summary>
        public bool Bold { get; set; }

        /// <summary>Gets or sets a value indicating whether the text is italic.</summary>
        public bool Italic { get; set; }

        /// <summary>Gets or sets a value indicating whether the text is underlined.</summary>
        public bool Underline { get; set; }

        /// <summary>Gets or sets a value indicating whether the text is struck through.</summary>
        public bool Strikethrough { get; set; }

        /// <summary>Gets or sets the font family name, or <see langword="null"/> for the default font.</summary>
        public string? FontName { get; set; }

        /// <summary>Gets or sets the font size in points; 0 for the default size.</summary>
        public double FontSize { get; set; }

        /// <summary>Gets or sets the font color as <c>0xRRGGBB</c>; -1 for the default color.</summary>
        public int FontColor { get; set; } = -1;

        /// <summary>Gets or sets the solid fill color as <c>0xRRGGBB</c>; -1 for no fill.</summary>
        public int FillColor { get; set; } = -1;

        /// <summary>Gets or sets the horizontal alignment.</summary>
        public ExcelHorizontalAlignment HorizontalAlignment { get; set; }

        /// <summary>Gets or sets the vertical alignment.</summary>
        public ExcelVerticalAlignment VerticalAlignment { get; set; }

        /// <summary>Gets or sets a value indicating whether text wraps within the cell.</summary>
        public bool WrapText { get; set; }

        /// <summary>Gets or sets the indent level (0 to 15).</summary>
        public int Indent { get; set; }

        /// <summary>Gets or sets the text rotation; see <see cref="CellStyle.Rotation"/>.</summary>
        public int Rotation { get; set; }

        /// <summary>Gets or sets the line style of all four border edges.</summary>
        public ExcelBorderStyle Border { get; set; }

        /// <summary>Gets or sets the border color as <c>0xRRGGBB</c>; -1 for automatic.</summary>
        public int BorderColor { get; set; } = -1;

        /// <summary>Converts the attribute's values to a <see cref="CellStyle"/>.</summary>
        /// <returns>The style.</returns>
        /// <exception cref="ArgumentException">A value is out of range.</exception>
        public CellStyle ToCellStyle()
        {
            CellBorder? border = null;
            if (Border != ExcelBorderStyle.None)
            {
                border = new CellBorder(Border, Color(BorderColor));
            }
            double? fontSize = null;
            if (FontSize != 0)
            {
                fontSize = FontSize;
            }
            var style = new CellStyle
            {
                NumberFormat = NumberFormat,
                Bold = Bold,
                Italic = Italic,
                Underline = Underline,
                Strikethrough = Strikethrough,
                FontName = FontName,
                FontSize = fontSize,
                FontColor = Color(FontColor),
                FillColor = Color(FillColor),
                HorizontalAlignment = HorizontalAlignment,
                VerticalAlignment = VerticalAlignment,
                WrapText = WrapText,
                Indent = Indent,
                Rotation = Rotation,
                Left = border,
                Right = border,
                Top = border,
                Bottom = border,
            };
            CellStyleValidation.Validate(in style, nameof(ExcelStyleAttribute));
            return style;
        }

        private static ExcelColor? Color(int rgb)
        {
            if (rgb == -1)
            {
                return null;
            }
            return ExcelColor.FromRgb(rgb);
        }
    }
}
