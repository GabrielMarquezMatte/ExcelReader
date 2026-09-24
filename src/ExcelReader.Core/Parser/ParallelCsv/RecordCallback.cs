using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Core.Parser.ParallelCsv
{
    internal static class RecordCallback<TRecord>
        where TRecord : ICsvRecord<TRecord>, allows ref struct
    {
        internal static CsvAggregation<byte> For(Action<TRecord> body)
        {
            return new CsvAggregation<byte>
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
