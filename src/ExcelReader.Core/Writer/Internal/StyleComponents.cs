using System.Runtime.InteropServices;

namespace ExcelReader.Core.Writer.Internal
{
    [StructLayout(LayoutKind.Auto)]
    internal readonly record struct FontKey(string? Name, double? Size, ExcelColor? Color, bool Bold, bool Italic, bool Underline, bool Strikethrough)
    {
        internal static FontKey Of(in CellStyle style)
            => new(style.FontName, style.FontSize, style.FontColor, style.Bold, style.Italic, style.Underline, style.Strikethrough);
    }

    internal readonly record struct FillKey(ExcelColor? Color);

    [StructLayout(LayoutKind.Auto)]
    internal readonly record struct BorderKey(CellBorder? Left, CellBorder? Right, CellBorder? Top, CellBorder? Bottom)
    {
        internal static BorderKey Of(in CellStyle style) => new(style.Left, style.Right, style.Top, style.Bottom);
    }

    internal sealed record StyleLimits(int MaxCellXfs, int MaxFonts, int MaxPaletteColors)
    {
        internal static readonly StyleLimits Ooxml = new(65_490, 512, int.MaxValue);

        internal static readonly StyleLimits Biff8 = new(4_000, 512, 56) { ReservedFontIndices = 1 };

        // BIFF8 skips font index 4, so that slot counts against MaxFonts.
        internal int ReservedFontIndices { get; init; }
    }

    // Fills 0 and 1 are the two mandatory built-ins (none, gray125) and appear as FillKey(null) placeholders.
    internal sealed class StyleComponents(IReadOnlyList<FontKey> fonts, IReadOnlyList<FillKey> fills, IReadOnlyList<BorderKey> borders,
        int[] fontIds, int[] fillIds, int[] borderIds)
    {
        internal IReadOnlyList<FontKey> Fonts { get; } = fonts;

        internal IReadOnlyList<FillKey> Fills { get; } = fills;

        internal IReadOnlyList<BorderKey> Borders { get; } = borders;

        internal int FontId(int styleIndex) => fontIds[styleIndex];

        internal int FillId(int styleIndex) => fillIds[styleIndex];

        internal int BorderId(int styleIndex) => borderIds[styleIndex];
    }
}
