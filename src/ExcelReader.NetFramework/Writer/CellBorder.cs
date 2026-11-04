namespace ExcelReader.NetFramework.Writer
{
    /// <summary>One edge of a cell border.</summary>
    public sealed class CellBorder
    {
        /// <summary>Creates a border edge.</summary>
        /// <param name="style">The line style.</param>
        /// <param name="color">The color as <c>0xRRGGBB</c>; -1 for automatic.</param>
        public CellBorder(ExcelBorderStyle style, int color = -1)
        {
            Style = style;
            Color = color;
        }

        /// <summary>Gets the line style.</summary>
        public ExcelBorderStyle Style { get; }

        /// <summary>Gets the color as <c>0xRRGGBB</c>; -1 for automatic.</summary>
        public int Color { get; }
    }
}
