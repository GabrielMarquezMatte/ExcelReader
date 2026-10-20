using ExcelReader.Core.Reader;
using ExcelReader.Native.Reading;

namespace ExcelReader.Native.Csv
{
    internal static unsafe partial class CsvAggregateApi
    {
        internal static int AggregateWorkbook(
            NativeHandle? handle, int sheet, NativeCsvAggregationRaw aggregation,
            int headerRow, int degreeOfParallelism, out nint result)
        {
            result = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            NativeApi.ClearLastError();
            if (headerRow < 0 || degreeOfParallelism < 0)
            {
                NativeApi.SetLastError(
                    $"header_row and degree_of_parallelism must not be negative; got {headerRow} and {degreeOfParallelism}.");
                return NativeStatus.InvalidArgument;
            }
            int status = ReadApi.ResolveSheet(handle, sheet, out IExcelSheet? resolved);
            if (status != NativeStatus.Ok)
            {
                return status;
            }

            CsvAggregateContext context = new(aggregation);
            SheetParallelOptions options = new() { DegreeOfParallelism = degreeOfParallelism, HeaderRow = headerRow };
            RowAggregation<CsvAggregateState> rows = Build(context, (ref CsvAggregateState state, Row row) =>
            {
                state.Views ??= new CellViewBuffer();
                int status;
                fixed (byte* rowValues = row.RowValues)
                fixed (byte* shared = row.Shared)
                fixed (byte* rowBuffer = row.RowBuffer)
                {
                    NativeCellView* cells = state.Views.Fill(row.Descriptors, rowValues, shared, rowBuffer);
                    status = context.AccumulateView((void*)state.Native, cells, row.Descriptors.Length, context.UserData);
                }
                if (status != NativeStatus.Ok)
                {
                    CsvAggregateContext.Abort(status);
                }
            });
            return Run(context, Task.Run(() => SheetParallel.AggregateAsync(resolved!, rows, options, CancellationToken.None)), out result);
        }
    }
}
