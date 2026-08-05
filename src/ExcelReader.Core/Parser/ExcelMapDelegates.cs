using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Reads a non-empty cell into a value of type <typeparamref name="TValue"/>. Every built-in reader
    /// exposed by <see cref="ExcelCellReaders"/> matches this shape, so a source-generated or
    /// hand-written <see cref="ExcelRowMapBuilder{T}"/> map can plug either a built-in reader or a
    /// custom one into <see cref="ExcelRowMapBuilder{T}.Property{TValue}"/>.
    /// </summary>
    /// <typeparam name="TValue">The value type the cell is read into.</typeparam>
    /// <param name="cell">The cell to read; callers only invoke this for a non-empty cell.</param>
    /// <param name="isDate1904">True when the source workbook uses the 1904 date system.</param>
    /// <param name="provider">The format provider configured for parsing.</param>
    /// <param name="value">The read value, when this method returns true.</param>
    /// <returns><see langword="true"/> if the cell was read successfully.</returns>
    public delegate bool ExcelCellReader<TValue>(in Cell cell, bool isDate1904, IFormatProvider provider, out TValue value);

    /// <summary>
    /// Assigns <paramref name="value"/> to a property of <paramref name="model"/>. The source generator
    /// emits this as a direct property assignment (e.g. <c>static (ref T m, V v) =&gt; m.Prop = v</c>) —
    /// no delegate compilation at runtime.
    /// </summary>
    /// <typeparam name="TModel">The row model type being assigned to.</typeparam>
    /// <typeparam name="TValue">The value type being assigned.</typeparam>
    /// <param name="model">The model instance to mutate.</param>
    /// <param name="value">The value to assign.</param>
    public delegate void ExcelPropertySetter<TModel, in TValue>(ref TModel model, TValue value)
#if NET9_0_OR_GREATER
        where TModel : allows ref struct
        where TValue : allows ref struct;
#else
        ;
#endif

    /// <summary>
    /// Reads a matched, non-empty cell and assigns the result directly onto <paramref name="model"/> in
    /// one call, instead of going through a separate <see cref="ExcelCellReader{TValue}"/> +
    /// <see cref="ExcelPropertySetter{TModel, TValue}"/> pair. Plug this into
    /// <see cref="ExcelRowMapBuilder{T}.PropertyRaw"/> when a single fused delegate (e.g. one emitted by
    /// the <c>[ExcelSerializable]</c> source generator) should replace what would otherwise be two
    /// indirect calls per bound cell per row.
    /// </summary>
    /// <typeparam name="TModel">The row model type being read into.</typeparam>
    /// <param name="model">The model instance to mutate.</param>
    /// <param name="cell">The cell to read; callers only invoke this for a non-empty cell.</param>
    /// <param name="isDate1904">True when the source workbook uses the 1904 date system.</param>
    /// <param name="provider">The format provider configured for parsing.</param>
    /// <returns><see langword="true"/> if the cell was read and assigned successfully.</returns>
    public delegate bool ExcelRowParser<TModel>(ref TModel model, in Cell cell, bool isDate1904, IFormatProvider provider)
#if NET9_0_OR_GREATER
        where TModel : allows ref struct;
#else
        ;
#endif
}
