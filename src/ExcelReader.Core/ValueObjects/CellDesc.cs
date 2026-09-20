using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using ExcelReader.Core.Enums;

namespace ExcelReader.Core.ValueObjects
{
    [StructLayout(LayoutKind.Auto)]
    internal readonly struct CellDesc
    {
        public int Column { get; init; }
        public int Start { get; init; }
        public int Length { get; init; }
        private readonly byte _type;
        public CellType Type { get => (CellType)_type; init => _type = (byte)value; }
        public int Style { get; init; }
        public CellValueSource Source { get; init; }
        public double Number { get; init; }
        public bool HasNumber { get; init; }
        public int SharedIndex { get; init; }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal Cell ToCell(ReadOnlySpan<byte> rowValues, ReadOnlySpan<byte> shared, ReadOnlySpan<byte> rowBuffer,
            string?[]? sharedStringCache = null, Utf8StringCache? contentCache = null)
        {
            if (Source == CellValueSource.Shared)
            {
                return new Cell(Type, shared.Slice(Start, Length), Number, HasNumber, Style, SharedIndex, sharedStringCache, contentCache);
            }
            ReadOnlySpan<byte> buf = Source == CellValueSource.RowBuffer ? rowBuffer : rowValues;
            return new Cell(Type, buf.Slice(Start, Length), Number, HasNumber, Style, sharedIndex: -1, sharedCache: null, contentCache);
        }
    }
}
