using System.Runtime.ExceptionServices;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Xlsb;

namespace ExcelReader.Core.Parser.SheetPipeline
{
    internal sealed class XlsbPipelineSource(XlsbWorkbook workbook, int index) : SheetPipelineSource
    {
        private bool _rowOpen;

        internal override object Owner => workbook;

        internal override ReaderLifetime Lifetime => workbook.Lifetime;

        internal override int MaxCellBytes => workbook.MaxCellBytes;

        internal override long UncompressedSize => workbook.SheetUncompressedSize(index);

        internal override PipelineStart? Begin(int headerRow, HeaderVisitor? onHeader)
        {
            long offset;
            using (XlsbWorkbook.Enumerator rows = workbook.OpenSheetWithoutPrefetch(index))
            {
                for (int row = 1; row <= headerRow; row++)
                {
                    if (!rows.MoveNext())
                    {
                        return null;
                    }
                    if (row == headerRow)
                    {
                        onHeader?.Invoke(rows.Current);
                    }
                }
                if (rows.Ended)
                {
                    return null;
                }
                _rowOpen = rows.RowOpen;
                offset = rows.Offset;
            }

            Stream stream = workbook.OpenSheetStream(index);
            try
            {
                Skip(stream, offset);
            }
            catch
            {
                stream.Dispose();
                throw;
            }
            return new PipelineStart(stream, new XlsbRowCutter());
        }

        internal override ChunkOutcome<TState> ParseChunk<TState>(
            ReadOnlyMemory<byte> chunk, bool first, RowAggregation<TState> aggregation, CancellationToken ct)
        {
            var outcome = new ChunkOutcome<TState>(aggregation.Seed());
            var rows = new XlsbWorkbook.Enumerator(workbook, chunk, first && _rowOpen, ct);
            try
            {
                RowAction<TState> accumulate = aggregation.Accumulate;
                while (rows.MoveNext())
                {
                    if (++outcome.Rows % RowsPerCancellationCheck == 0)
                    {
                        ct.ThrowIfCancellationRequested();
                    }
                    accumulate(ref outcome.State, rows.Current);
                }
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                outcome.Failure = ExceptionDispatchInfo.Capture(ex);
            }
            finally
            {
                rows.Dispose();
            }
            return outcome;
        }

        internal override ChunkOutcome<TState> ParseTail<TState>(int headerRow, long skipRows, RowAggregation<TState> aggregation, CancellationToken ct)
        {
            throw new System.Diagnostics.UnreachableException("An XLSB chunk always ends on a record boundary.");
        }
    }
}
