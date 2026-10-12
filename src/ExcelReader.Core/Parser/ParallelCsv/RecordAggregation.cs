using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal static class RecordAggregation<TAccumulator, TRecord>
        where TAccumulator : IRowAccumulator<TAccumulator, TRecord>, new()
        where TRecord : IRowRecord<TRecord>, allows ref struct
    {
        internal static readonly RowAggregation<TAccumulator> Instance = new()
        {
            Seed = static () => new TAccumulator(),
            Accumulate = static (ref accumulator, row) =>
            {
                if (TRecord.TryParse(row, out TRecord record))
                {
                    accumulator.Add(record);
                }
            },
            Combine = static (left, right) =>
            {
                left.Merge(right);
                return left;
            },
        };
    }
}
