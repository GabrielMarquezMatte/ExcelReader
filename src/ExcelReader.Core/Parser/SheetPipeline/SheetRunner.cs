using System.Runtime.CompilerServices;
using ExcelReader.Core.Parser.ParallelCsv;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Xls;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Core.Parser.SheetPipeline
{
    internal static class SheetRunner
    {
        private const int RowsPerCancellationCheck = 4096;

        internal static Task<TState> RunAsync<TState>(
            IExcelSheet sheet, RowAggregation<TState> aggregation, RowActionFactory<TState>? factory, SheetParallelOptions options, CancellationToken ct)
        {
            if (sheet is CsvSheet csv)
            {
                return ParallelCsvProcessor.RunAsync(csv.Reader, aggregation, factory, options, ct);
            }
            int dop = ParallelCsvFactory.Normalize(options.DegreeOfParallelism);
            SheetPipelineSource? pipeline = SheetPipelineSource.TryCreate(sheet, dop);
            if (pipeline is not null)
            {
                return SheetPipelineRunner.AggregateAsync(pipeline, aggregation, factory, options.HeaderRow, dop, SheetPipelineSource.DefaultChunkBytes, ct);
            }
            return SequentialAsync(sheet, aggregation, factory, options.HeaderRow, ct);
        }

        internal static IAsyncEnumerable<T> Parse<T>(IExcelSheet sheet, ExcelParser<T> parser, int degreeOfParallelism, CancellationToken ct)
        {
            if (sheet is CsvSheet csv)
            {
                return ParallelCsvFactory.Create(csv.Reader, degreeOfParallelism, parser, ct);
            }
            int dop = ParallelCsvFactory.Normalize(degreeOfParallelism);
            SheetPipelineSource? pipeline = SheetPipelineSource.TryCreate(sheet, dop);
            if (pipeline is not null)
            {
                return SheetRowProjection<T>.ParseAsync(pipeline, parser, sheet.IsDate1904, dop, SheetPipelineSource.DefaultChunkBytes, ct);
            }
            return Sequential(sheet, parser, ct);
        }

        private static async IAsyncEnumerable<T> Sequential<T>(IExcelSheet sheet, ExcelParser<T> parser, [EnumeratorCancellation] CancellationToken ct)
        {
            IAsyncEnumerable<T> rows = sheet switch
            {
                XlsxSheet xlsx => parser.Parse(xlsx),
                XlsbSheet xlsb => parser.Parse(xlsb),
                XlsSheet xls => parser.Parse(xls),
                _ => parser.Parse(sheet),
            };
            await foreach (T row in rows.WithCancellation(ct).ConfigureAwait(false))
            {
                yield return row;
            }
        }

        internal static async Task<TState> SequentialAsync<TState>(
            IExcelSheet sheet, RowAggregation<TState> aggregation, RowActionFactory<TState>? factory, int headerRow, CancellationToken ct)
        {
            IExcelRowEnumerator rows = sheet.GetAsyncEnumerator(ct);
            await using (rows.ConfigureAwait(false))
            {
                TState state = aggregation.Seed();
                RowAction<TState> accumulate = aggregation.Accumulate;
                if (factory is not null && headerRow == 0)
                {
                    accumulate = factory(default, sequential: true);
                }
                for (int skipped = 0; skipped < headerRow; skipped++)
                {
                    if (!await rows.MoveNextAsync().ConfigureAwait(false))
                    {
                        return state;
                    }
                    if (factory is not null && skipped == headerRow - 1)
                    {
                        accumulate = factory(rows.Current, sequential: true);
                    }
                }
                long delivered = 0;
                while (await rows.MoveNextAsync().ConfigureAwait(false))
                {
                    if (++delivered % RowsPerCancellationCheck == 0)
                    {
                        ct.ThrowIfCancellationRequested();
                    }
                    Deliver(accumulate, ref state, rows);
                }
                return state;
            }
        }

        private static void Deliver<TState>(RowAction<TState> accumulate, ref TState state, IExcelRowEnumerator rows)
        {
            accumulate(ref state, rows.Current);
        }
    }
}
