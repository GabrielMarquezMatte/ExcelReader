using System.Runtime.InteropServices;
using ExcelReader.Core.Enums;

namespace ExcelReader.Core.ValueObjects
{
    // One parsed cell, located in one of Row's three byte spans (rowValues / shared / rowBuffer).
    [StructLayout(LayoutKind.Auto)]
    internal readonly struct CellDesc
    {
        public int Column { get; init; }
        public int Start { get; init; }
        public int Length { get; init; }
        public CellType Type { get; init; }
        public int Style { get; init; }
        public CellValueSource Source { get; init; }
        // Raw numeric value, set when the source stored a binary double (XLS Number/RK/Date/Formula).
        // Lets consumers skip the format-on-read / parse-on-consume round trip. Start/Length still
        // point at the formatted text, so Value/GetString stay byte-identical.
        public double Number { get; init; }
        public bool HasNumber { get; init; }

        internal Cell ToCell(ReadOnlySpan<byte> rowValues, ReadOnlySpan<byte> shared, ReadOnlySpan<byte> rowBuffer, Dictionary<int, string>? sharedStringCache = null)
        {
            // `Start` is only usable as a dedup-cache key when it indexes a stable, append-only,
            // cross-row buffer (a true shared-string table) — reject the cache otherwise, since CSV
            // reuses Source.Shared/Start for its own per-row materialized scratch buffer, where the same
            // Start is reused by unrelated content on the next row. RowBuffer cells (aliasing the live
            // read buffer directly) are even less stable across rows, so they must never reach the cache.
            if (Source == CellValueSource.Shared)
            {
                return new Cell(Type, shared.Slice(Start, Length), Number, HasNumber, Style, Start, sharedStringCache);
            }
            ReadOnlySpan<byte> buf = Source == CellValueSource.RowBuffer ? rowBuffer : rowValues;
            return new Cell(Type, buf.Slice(Start, Length), Number, HasNumber, Style);
        }
    }
}
