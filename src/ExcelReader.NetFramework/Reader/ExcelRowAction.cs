namespace ExcelReader.NetFramework.Reader
{
    /// <summary>Receives one row of <see cref="ExcelSheet.ForEach"/>.</summary>
    /// <param name="row">The row, valid only for the duration of the call.</param>
    public delegate void ExcelRowAction(ExcelRowView row);
}
