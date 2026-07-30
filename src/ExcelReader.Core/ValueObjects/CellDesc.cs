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
        // Shared-string table index (not a byte offset) — the cache key into the reader's per-index
        // string?[] dedup cache. -1 for non-Shared cells and for an out-of-range/corrupt shared index
        // (see WorkbookLookups.SharedAt), which Cell.GetString() treats as "no cache entry".
        public int SharedIndex { get; init; }

        internal Cell ToCell(ReadOnlySpan<byte> rowValues, ReadOnlySpan<byte> shared, ReadOnlySpan<byte> rowBuffer, string?[]? sharedStringCache = null)
        {
            // SharedIndex is only usable as a dedup-cache key when it indexes a stable, append-only,
            // cross-row shared-string table (XLSX/XLSB/XLS) — CSV reuses Source.Shared/Start for its own
            // per-row materialized scratch buffer and never wires a cache (Row's 3-arg constructors leave
            // sharedStringCache null), so this branch is unreachable for CSV regardless. RowBuffer cells
            // (aliasing the live read buffer directly) are even less stable across rows and never use
            // Source.Shared, so they always fall to the plain slice below.
            if (Source == CellValueSource.Shared)
            {
                return new Cell(Type, shared.Slice(Start, Length), Number, HasNumber, Style, SharedIndex, sharedStringCache);
            }
            ReadOnlySpan<byte> buf = Source == CellValueSource.RowBuffer ? rowBuffer : rowValues;
            return new Cell(Type, buf.Slice(Start, Length), Number, HasNumber, Style);
        }
    }
}
