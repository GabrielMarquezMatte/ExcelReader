using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    // Setter compiled once per property via Expression tree.
    // ref TModel allows in-place mutation for both classes and structs.
    internal delegate void RefAction<TModel, in TProperty>(ref TModel model, TProperty value)
#if NET9_0_OR_GREATER
        where TModel : allows ref struct;
#else
        ;
#endif

    // Column-level TryParse over a cell already matched to the target column.
    // Returns false on parse failure; true on success or empty cell (keep default).
    // provider supplies the culture for text-backed numeric/Guid cells (ExcelParserConfig.Culture).
    internal delegate bool ColumnParser<TModel>(
        ref TModel model,
        in Cell cell,
        bool isDate1904,
        IFormatProvider provider)
#if NET9_0_OR_GREATER
        where TModel : allows ref struct;
#else
        ;
#endif
}
