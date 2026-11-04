namespace ExcelReader.NetFramework.Writer
{
    /// <summary>A cell style: number format, font, fill, borders and alignment. Register one with <see cref="ExcelWorkbookWriter.AddStyle"/>. Colors are <c>0xRRGGBB</c>, -1 for unset.</summary>
    public sealed class CellStyle
    {
        /// <summary>Gets or sets the Excel number-format code (e.g. <c>"#,##0.00"</c>), or null for the general format.</summary>
        public string? NumberFormat { get; set; }

        /// <summary>Gets or sets a value indicating whether the text is bold.</summary>
        public bool Bold { get; set; }

        /// <summary>Gets or sets a value indicating whether the text is italic.</summary>
        public bool Italic { get; set; }

        /// <summary>Gets or sets a value indicating whether the text is underlined.</summary>
        public bool Underline { get; set; }

        /// <summary>Gets or sets a value indicating whether the text is struck through.</summary>
        public bool Strikethrough { get; set; }

        /// <summary>Gets or sets the font family name (1 to 31 characters), or null for the default font.</summary>
        public string? FontName { get; set; }

        /// <summary>Gets or sets the font size in points (1 to 409), or null for the default size.</summary>
        public double? FontSize { get; set; }

        /// <summary>Gets or sets the font color; -1 for the default color.</summary>
        public int FontColor { get; set; } = -1;

        /// <summary>Gets or sets the solid background color; -1 for no fill.</summary>
        public int FillColor { get; set; } = -1;

        /// <summary>Gets or sets the left border, or null for none.</summary>
        public CellBorder? Left { get; set; }

        /// <summary>Gets or sets the right border, or null for none.</summary>
        public CellBorder? Right { get; set; }

        /// <summary>Gets or sets the top border, or null for none.</summary>
        public CellBorder? Top { get; set; }

        /// <summary>Gets or sets the bottom border, or null for none.</summary>
        public CellBorder? Bottom { get; set; }

        /// <summary>Gets or sets the horizontal alignment.</summary>
        public ExcelHorizontalAlignment HorizontalAlignment { get; set; }

        /// <summary>Gets or sets the vertical alignment.</summary>
        public ExcelVerticalAlignment VerticalAlignment { get; set; }

        /// <summary>Gets or sets a value indicating whether text wraps within the cell.</summary>
        public bool WrapText { get; set; }

        /// <summary>Gets or sets the indent level (0 to 15).</summary>
        public int Indent { get; set; }

        /// <summary>Gets or sets the text rotation in degrees (-90 to 90), or 255 for stacked text.</summary>
        public int Rotation { get; set; }
    }
}
