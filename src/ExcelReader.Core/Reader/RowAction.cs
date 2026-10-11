namespace ExcelReader.Core.Reader
{
    /// <summary>Processes one row into the state owned by the partition that read it.</summary>
    /// <typeparam name="TState">The per-partition accumulator type.</typeparam>
    /// <param name="state">The accumulator of the partition the row belongs to. Mutate only this.</param>
    /// <param name="row">The row. Valid only for the duration of the call.</param>
    public delegate void RowAction<TState>(ref TState state, Row row);
}
