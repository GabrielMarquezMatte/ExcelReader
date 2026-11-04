using System.Runtime.InteropServices;

namespace ExcelReader.Core.Writer
{
    /// <summary>An opaque RGB color.</summary>
    /// <param name="R">The red component.</param>
    /// <param name="G">The green component.</param>
    /// <param name="B">The blue component.</param>
    [StructLayout(LayoutKind.Auto)]
    public readonly record struct ExcelColor(byte R, byte G, byte B)
    {
        /// <summary>Creates a color from a <c>0xRRGGBB</c> value.</summary>
        /// <param name="rgb">The packed color, from 0 to 0xFFFFFF.</param>
        /// <returns>The color.</returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="rgb"/> is outside 0 to 0xFFFFFF.</exception>
        public static ExcelColor FromRgb(int rgb)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(rgb);
            ArgumentOutOfRangeException.ThrowIfGreaterThan(rgb, 0xFFFFFF);
            return new ExcelColor((byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb);
        }

        internal int ToRgb()
        {
            return (R << 16) | (G << 8) | B;
        }
    }
}
