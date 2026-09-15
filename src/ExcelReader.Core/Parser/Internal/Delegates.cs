using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    // Setter compiled once per property via Expression tree.
    // ref TModel allows in-place mutation for both classes and structs.
    // TProperty also allows ref struct so a ReadOnlySpan<byte> property can bind directly to
    // Cell.Value (zero-copy) via ColumnParserFactory's span parser — see BuildSpanParser.
    internal delegate void RefAction<TModel, in TProperty>(ref TModel model, TProperty value)
        where TModel : allows ref struct
        where TProperty : allows ref struct;

    // Column-level TryParse over a cell already matched to the target column. Callers never invoke
    // this for an empty cell (they skip the call and keep the model's default), so implementations
    // don't need their own empty-cell guard. Returns false on parse failure, true on success.
    // provider supplies the culture for text-backed numeric/Guid cells (ExcelParserConfig.Culture).
    internal delegate bool ColumnParser<TModel>(
        ref TModel model,
        in Cell cell,
        bool isDate1904,
        IFormatProvider provider)
        where TModel : allows ref struct;
}
