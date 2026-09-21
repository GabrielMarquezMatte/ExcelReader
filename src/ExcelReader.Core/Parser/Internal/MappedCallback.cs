using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser.Internal
{
    internal static class MappedCallback<TModel>
        where TModel : allows ref struct
    {
        internal static readonly CsvAggregation<byte> Unbound = new()
        {
            Seed = static () => 0,
            Accumulate = static (ref _, _) => throw new InvalidOperationException("The model map was not bound before the first record."),
            Combine = static (left, _) => left,
        };

        internal static CsvAccumulateFactory<byte> Binder(CsvModelMap<TModel> map, int headerRow, CsvRecordAction<TModel> body)
        {
            CsvModelSink<byte, TModel> sink = (ref byte _, TModel model) => body(model);
            return (Row header, bool sequential) =>
                MappedProjection<byte, TModel>.Bind(map, header, headerRow, sequential, sink);
        }
    }
}
