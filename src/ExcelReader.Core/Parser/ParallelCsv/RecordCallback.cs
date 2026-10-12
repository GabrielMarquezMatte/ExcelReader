using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal static class RecordCallback<TRecord>
        where TRecord : IRowRecord<TRecord>, allows ref struct
    {
        internal static RowAggregation<byte> For(Action<TRecord> body)
        {
            return new RowAggregation<byte>
            {
                Seed = static () => 0,
                Accumulate = (ref byte _, Row row) =>
                {
                    if (TRecord.TryParse(row, out TRecord record))
                    {
                        body(record);
                    }
                },
                Combine = static (left, _) => left,
            };
        }
    }
}
