using System.Runtime.InteropServices;
using ExcelReader.Core.Reader.Schema;

namespace ExcelReader.Core.Reader.Columnar
{
    internal readonly struct TypedColumnSpec
    {
        public TypedColumnSpec()
        {
        }

        internal string[] Names { get; init; } = [];
        internal int Index { get; init; }
        internal ExcelColumnType Type { get; init; }
        internal bool Nullable { get; init; }

        /// <summary>The name a conversion failure reports when <see cref="Names"/> is empty.</summary>
        internal string? Label { get; init; }
    }

    /// <summary>
    /// One finished column as native blocks. Strings keep their offsets in <see cref="Values"/> and
    /// their bytes in <see cref="Data"/>; a dictionary column keeps its int32 indices in
    /// <see cref="Values"/> and its distinct strings in <see cref="DictionaryOffsets"/> and
    /// <see cref="Data"/>. A consumer that takes a block zeroes its field.
    /// </summary>
    [StructLayout(LayoutKind.Auto)]
    internal struct BuiltColumn
    {
        public ExcelColumnType Type;
        public long Length;
        public IntPtr Values;
        public IntPtr Validity;
        public IntPtr Data;
        public long DataLen;
        public IntPtr DictionaryOffsets;
        public long DictionaryLength;

        internal void Free()
        {
            ColumnBlocks.Free(Values);
            ColumnBlocks.Free(Validity);
            ColumnBlocks.Free(Data);
            ColumnBlocks.Free(DictionaryOffsets);
            this = default;
        }
    }

    internal sealed class ColumnarTable(long rowCount, BuiltColumn[] columns) : IDisposable
    {
        internal long RowCount { get; } = rowCount;

        internal BuiltColumn[] Columns { get; } = columns;

        public void Dispose()
        {
            for (int i = 0; i < Columns.Length; i++)
            {
                Columns[i].Free();
            }
        }
    }
}
