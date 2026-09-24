namespace ExcelReader.Core.Reader.Internal
{
    internal enum CellValueSource : byte
    {
        RowValues,
        Shared,
        RowBuffer,
    }
}
