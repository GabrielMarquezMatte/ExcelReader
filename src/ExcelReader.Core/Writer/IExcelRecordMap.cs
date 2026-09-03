namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// Implemented by the source generator (feature A) — or by hand, before the generator runs — on the
    /// record type itself, which must be <c>partial</c>. Generated in the same <c>partial</c> as
    /// <see cref="Parser.IExcelRowMap{T}"/>, since both maps derive from the same
    /// property list and attributes.
    /// </summary>
    /// <typeparam name="T">The record type this map configures.</typeparam>
    public interface IExcelRecordMap<T>
    {
        /// <summary>Configures <paramref name="builder"/> with one column per mapped property.</summary>
        /// <typeparam name="TRow">
        /// The concrete row writer the columns write through; generic so each column action binds to
        /// that sealed class instead of <see cref="IRowWriter"/> (see
        /// <see cref="ExcelRecordMapBuilder{T, TRow}"/>). An implementation's body is identical for
        /// every <typeparamref name="TRow"/> — write it once as a generic method.
        /// </typeparam>
        /// <param name="builder">The builder to configure.</param>
        static abstract void ConfigureExcelRecordMap<TRow>(ExcelRecordMapBuilder<T, TRow> builder)
            where TRow : IRowWriter;
    }
}
