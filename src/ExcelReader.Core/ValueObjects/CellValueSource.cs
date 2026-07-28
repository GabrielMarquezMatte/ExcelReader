namespace ExcelReader.Core.ValueObjects
{
    // Which of Row's three byte spans a CellDesc's Start/Length index into. RowBuffer is distinct from
    // RowValues so the shared-string dedup cache (keyed on Start, see CellDesc.ToCell) never sees a
    // RowBuffer cell: RowBuffer offsets are only stable for the current row, unlike a true shared table.
    internal enum CellValueSource : byte
    {
        RowValues,
        Shared,
        RowBuffer,
    }
}
