using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal static class RecordAggregation<TAccumulator, TRecord>
        where TAccumulator : ICsvAccumulator<TAccumulator, TRecord>, new()
        where TRecord : ICsvRecord<TRecord>, allows ref struct
    {
        internal static readonly CsvAggregation<TAccumulator> Instance = new()
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
