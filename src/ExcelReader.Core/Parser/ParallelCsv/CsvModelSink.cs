namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal delegate void CsvModelSink<TState, TModel>(ref TState state, TModel model)
        where TModel : allows ref struct;
}
