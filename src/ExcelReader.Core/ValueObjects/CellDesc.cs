using System.Runtime.InteropServices;
using ExcelReader.Core.Enums;

namespace ExcelReader.Core.ValueObjects
{
    // One parsed cell, located in either the shared-strings flat buffer or the per-row value buffer.
    [StructLayout(LayoutKind.Auto)]
    internal readonly struct CellDesc
    {
        public int Column { get; init; }
        public int Start { get; init; }
        public int Length { get; init; }
        public CellType Type { get; init; }
        public int Style { get; init; }
        public bool FromShared { get; init; }
        // Raw numeric value, set when the source stored a binary double (XLS Number/RK/Date/Formula).
        // Lets consumers skip the format-on-read / parse-on-consume round trip. Start/Length still
        // point at the formatted text, so Value/GetString stay byte-identical.
        public double Number { get; init; }
        public bool HasNumber { get; init; }

        internal Cell ToCell(ReadOnlySpan<byte> rowValues, ReadOnlySpan<byte> shared)
        {
            var buf = FromShared ? shared : rowValues;
            return new Cell(Type, buf.Slice(Start, Length), Number, HasNumber, Style);
        }
    }
}
