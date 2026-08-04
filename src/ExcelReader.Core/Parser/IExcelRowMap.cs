namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Implemented by the source generator (feature A) — or by hand, before the generator runs — on the
    /// row model type itself, which must be <c>partial</c>. The <see langword="static abstract"/> member
    /// turns "no map was generated for this model" into a compile error (<typeparamref name="T"/> fails
    /// the <c>where T : IExcelRowMap&lt;T&gt;</c> constraint on <see cref="ExcelMappedParser{T}"/>)
    /// instead of a silent runtime fallback to reflection.
    /// </summary>
    /// <typeparam name="T">The row model type this map configures.</typeparam>
    public interface IExcelRowMap<T>
#if NET9_0_OR_GREATER
        where T : allows ref struct
#endif
    {
        /// <summary>Configures <paramref name="builder"/> with one binding per mapped property.</summary>
        /// <param name="builder">The builder to configure.</param>
        static abstract void ConfigureExcelRowMap(ExcelRowMapBuilder<T> builder);
    }
}
