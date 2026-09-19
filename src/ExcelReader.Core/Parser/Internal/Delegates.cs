using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    internal delegate void RefAction<TModel, in TProperty>(ref TModel model, TProperty value)
        where TModel : allows ref struct
        where TProperty : allows ref struct;

    internal delegate bool ColumnParser<TModel>(
        ref TModel model,
        in Cell cell,
        bool isDate1904,
        IFormatProvider provider)
        where TModel : allows ref struct;
}
