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
    }
}
