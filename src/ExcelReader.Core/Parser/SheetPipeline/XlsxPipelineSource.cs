using System.Runtime.ExceptionServices;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Core.Parser.SheetPipeline
{
    internal sealed class XlsxPipelineSource(XlsxWorkbook workbook, int index) : SheetPipelineSource
    {
        private NsTokens? _ns;

        internal override object Owner => workbook;

        internal override ReaderLifetime Lifetime => workbook.Lifetime;

        internal override int MaxCellBytes => workbook.MaxCellBytes;

        internal override long UncompressedSize => workbook.SheetUncompressedSize(index);

        internal override PipelineStart? Begin(int headerRow, HeaderVisitor? onHeader)
        {
            long offset;
            using (XlsxWorkbook.Enumerator rows = workbook.OpenSheetWithoutPrefetch(index))
            {
                _ns = rows.Namespace;
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
            return new PipelineStart(stream, new XlsxRowCutter(_ns?.RowOpen ?? "<row"u8.ToArray(), _ns?.SheetDataEnd ?? "</sheetData"u8.ToArray()));
        }

        internal override ChunkOutcome<TState> ParseChunk<TState>(
            ReadOnlyMemory<byte> chunk, bool first, RowAggregation<TState> aggregation, CancellationToken ct)
        {
            var outcome = new ChunkOutcome<TState>(aggregation.Seed());
            var rows = new XlsxWorkbook.Enumerator(workbook, chunk, _ns, ct);
            try
            {
                RowAction<TState> accumulate = aggregation.Accumulate;
                while (rows.MoveNext() && !rows.EndedMidMarkup)
                {
                    if (++outcome.Rows % RowsPerCancellationCheck == 0)
                    {
                        ct.ThrowIfCancellationRequested();
                    }
                    accumulate(ref outcome.State, rows.Current);
                }
                outcome.Clean = !rows.EndedMidMarkup;
                outcome.ReachedEnd = rows.ReachedEnd;
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
            var outcome = new ChunkOutcome<TState>(aggregation.Seed());
            using XlsxWorkbook.Enumerator rows = workbook.OpenSheetWithoutPrefetch(index);
            for (long skipped = 0; skipped < headerRow + skipRows; skipped++)
            {
                if (!rows.MoveNext())
                {
                    return outcome;
                }
            }
            RowAction<TState> accumulate = aggregation.Accumulate;
            try
            {
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
            return outcome;
        }
    }
}
