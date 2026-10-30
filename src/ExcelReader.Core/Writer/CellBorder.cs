using System.Runtime.InteropServices;

namespace ExcelReader.Core.Writer
{
    /// <summary>One edge of a cell border.</summary>
    /// <param name="Style">The line style.</param>
    /// <param name="Color">The line color, or <see langword="null"/> for automatic.</param>
    [StructLayout(LayoutKind.Auto)]
    public readonly record struct CellBorder(ExcelBorderStyle Style, ExcelColor? Color = null);
}
