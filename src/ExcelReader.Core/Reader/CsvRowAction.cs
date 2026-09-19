using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
{
    /// <summary>Processes one CSV record into the state owned by the partition that read it.</summary>
    /// <typeparam name="TState">The per-partition accumulator type.</typeparam>
    /// <param name="state">The accumulator of the partition the record belongs to. Mutate only this.</param>
    /// <param name="row">The record. Valid only for the duration of the call.</param>
    public delegate void CsvRowAction<TState>(ref TState state, Row row);
}
