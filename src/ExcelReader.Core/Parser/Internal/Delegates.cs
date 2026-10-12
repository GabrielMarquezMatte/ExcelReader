using ExcelReader.Core.Reader;


namespace ExcelReader.Core.Parser.Internal
{
    internal delegate void RefAction<TModel, in TProperty>(ref TModel model, TProperty value)
        where TModel : allows ref struct
        where TProperty : allows ref struct;

    internal delegate bool ColumnParser<TModel>(
        ref TModel model,
        in Cell cell,
        in ExcelRowContext context)
        where TModel : allows ref struct;
}
