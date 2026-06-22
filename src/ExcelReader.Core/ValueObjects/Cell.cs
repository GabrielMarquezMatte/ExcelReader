using System.Diagnostics.CodeAnalysis;
using System.Text;
using ExcelReader.Core.Enums;

namespace ExcelReader.Core.ValueObjects
{
    public readonly ref struct Cell
    {
        public CellType Type { get; }
        // UTF-8 bytes of the cell's text: resolved shared-string text, or the raw <v> for numbers/bools/etc.
        public ReadOnlySpan<byte> Value { get; }
        // The `s` style index from the worksheet (0 when absent). Escape hatch for date detection,
        // which is deferred: dates arrive as Number, and the caller maps StyleIndex -> format itself.
        public int StyleIndex { get; }

        public Cell(CellType type, ReadOnlySpan<byte> value, int styleIndex = 0)
        {
            Type = type;
            Value = value;
            StyleIndex = styleIndex;
        }

        public bool TryParse<T>(IFormatProvider? provider, [MaybeNullWhen(false)] out T result) where T : IUtf8SpanParsable<T>
        {
            return T.TryParse(Value, provider, out result);
        }

        // Allocates — only call when you actually need a string.
        public string GetString()
        {
            return Encoding.UTF8.GetString(Value);
        }


        public override string ToString()
        {
            return GetString();
        }

    }
}
