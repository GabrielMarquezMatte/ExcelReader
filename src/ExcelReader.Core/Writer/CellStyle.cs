using System.Runtime.InteropServices;

namespace ExcelReader.Core.Writer
{
    /// <summary>A column- or row-level cell style: number format, font, fill, borders, and alignment. Register one with <see cref="IWorkbookWriter{TSheet}.AddStyle"/>.</summary>
    [StructLayout(LayoutKind.Auto)]
    public readonly record struct CellStyle
    {
        /// <summary>The <see cref="Rotation"/> value that stacks text vertically, one character above the next.</summary>
        public const int StackedTextRotation = 255;

        /// <summary>Gets the Excel number-format code (e.g. <c>"#,##0.00"</c>, <c>"yyyy-mm-dd"</c>, <c>"0.0%"</c>), or <see langword="null"/> for the general format.</summary>
        public string? NumberFormat { get; init; }

        /// <summary>Gets a value indicating whether text in cells using this style is bold.</summary>
        public bool Bold { get; init; }

        /// <summary>Gets a value indicating whether text in cells using this style is italic.</summary>
        public bool Italic { get; init; }

        /// <summary>Gets the font family name (1 to 31 characters), or <see langword="null"/> for the default font.</summary>
        public string? FontName { get; init; }

        /// <summary>Gets the font size in points (1 to 409), or <see langword="null"/> for the default size.</summary>
        public double? FontSize { get; init; }

        /// <summary>Gets the font color, or <see langword="null"/> for the default color.</summary>
        public ExcelColor? FontColor { get; init; }

        /// <summary>Gets a value indicating whether text in cells using this style is underlined.</summary>
        public bool Underline { get; init; }

        /// <summary>Gets a value indicating whether text in cells using this style is struck through.</summary>
        public bool Strikethrough { get; init; }

        /// <summary>Gets the solid background color, or <see langword="null"/> for no fill.</summary>
        public ExcelColor? FillColor { get; init; }

        /// <summary>Gets the left border, or <see langword="null"/> for none.</summary>
        public CellBorder? Left { get; init; }

        /// <summary>Gets the right border, or <see langword="null"/> for none.</summary>
        public CellBorder? Right { get; init; }

        /// <summary>Gets the top border, or <see langword="null"/> for none.</summary>
        public CellBorder? Top { get; init; }

        /// <summary>Gets the bottom border, or <see langword="null"/> for none.</summary>
        public CellBorder? Bottom { get; init; }

        /// <summary>Gets the horizontal alignment.</summary>
        public ExcelHorizontalAlignment HorizontalAlignment { get; init; }

        /// <summary>Gets the vertical alignment.</summary>
        public ExcelVerticalAlignment VerticalAlignment { get; init; }

        /// <summary>Gets a value indicating whether text wraps within the cell.</summary>
        public bool WrapText { get; init; }

        /// <summary>Gets the indent level (0 to 15).</summary>
        public int Indent { get; init; }

        /// <summary>Gets the text rotation in degrees (-90 to 90), or <see cref="StackedTextRotation"/>.</summary>
        public int Rotation { get; init; }
    }
}
