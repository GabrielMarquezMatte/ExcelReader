using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    // Setter compiled once per property via Expression tree.
    // ref TModel allows in-place mutation for both classes and structs.
    internal delegate void RefAction<TModel, in TProperty>(ref TModel model, TProperty value)
        where TModel : allows ref struct;

    // Column-level TryParse: accesses row[columnIndex] internally.
    // Returns false on parse failure; true on success or empty cell (keep default).
    internal delegate bool ColumnParser<TModel>(
        ref TModel model,
        in Row row,
        int columnIndex,
        bool isDate1904)
        where TModel : allows ref struct;
}
