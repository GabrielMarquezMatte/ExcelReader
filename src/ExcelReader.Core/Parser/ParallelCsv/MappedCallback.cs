using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal static class MappedCallback<TModel>
        where TModel : allows ref struct
    {
        internal static readonly RowAggregation<byte> Unbound = new()
        {
            Seed = static () => 0,
            Accumulate = static (ref _, _) => throw new InvalidOperationException("The model map was not bound before the first record."),
            Combine = static (left, _) => left,
        };

        internal static RowActionFactory<byte> Binder(ExcelParser<TModel> parser, int headerRow, Action<TModel> body)
        {
            ModelSink<byte, TModel> sink = (ref byte _, TModel model) => body(model);
            return (Row header, bool sequential) =>
                MappedProjection<byte, TModel>.Bind(parser, header, headerRow, sequential, sink);
        }
    }
}
