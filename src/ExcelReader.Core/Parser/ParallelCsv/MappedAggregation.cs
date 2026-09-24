using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal static class MappedAggregation<TAccumulator, TModel>
        where TAccumulator : ICsvAccumulator<TAccumulator, TModel>, new()
        where TModel : allows ref struct
    {
        private static readonly CsvModelSink<TAccumulator, TModel> Sink =
            static (ref TAccumulator accumulator, TModel model) => accumulator.Add(model);

        internal static readonly CsvAggregation<TAccumulator> Unbound = new()
        {
            Seed = static () => new TAccumulator(),
            Accumulate = static (ref _, _) => throw new InvalidOperationException("The model map was not bound before the first record."),
            Combine = static (left, right) =>
            {
                left.Merge(right);
                return left;
            },
        };

        internal static CsvAccumulateFactory<TAccumulator> Binder(ExcelParser<TModel> parser, int headerRow)
        {
            return (Row header, bool sequential) =>
                MappedProjection<TAccumulator, TModel>.Bind(parser, header, headerRow, sequential, Sink);
        }
    }
}
