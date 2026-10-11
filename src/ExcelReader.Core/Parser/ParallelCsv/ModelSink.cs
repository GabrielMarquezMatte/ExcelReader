namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal delegate void ModelSink<TState, TModel>(ref TState state, TModel model)
        where TModel : allows ref struct;
}
