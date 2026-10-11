using System.Buffers;
using System.Runtime.ExceptionServices;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Core.Parser.SheetPipeline
{
    internal delegate void HeaderVisitor(Row header);

    internal sealed record PipelineStart(Stream Stream, RowCutter Cutter);

    internal sealed class ChunkOutcome<TState>(TState state)
    {
        public TState State = state;
        public long Rows;

        // False when the chunk did not end between rows: its rows are valid, but what follows must be read sequentially.
        public bool Clean = true;
        public bool ReachedEnd;
        public ExceptionDispatchInfo? Failure;
    }

    internal abstract class SheetPipelineSource
    {
        internal const long MinBytes = 256 * 1024;
        internal const int DefaultChunkBytes = 1024 * 1024;
        private protected const int RowsPerCancellationCheck = 4096;

        internal abstract object Owner { get; }

        internal abstract ReaderLifetime Lifetime { get; }

        internal abstract int MaxCellBytes { get; }

        internal abstract long UncompressedSize { get; }

        // Null when the sheet ends at or before headerRow; otherwise the stream is positioned at the first data row.
        internal abstract PipelineStart? Begin(int headerRow, HeaderVisitor? onHeader);

        internal abstract ChunkOutcome<TState> ParseChunk<TState>(
            ReadOnlyMemory<byte> chunk, bool first, RowAggregation<TState> aggregation, CancellationToken ct);

        // Reads the sheet again from its start: skipRows counts data rows, after the header.
        internal abstract TState ParseTail<TState>(int headerRow, long skipRows, RowAggregation<TState> aggregation, CancellationToken ct);

        internal static SheetPipelineSource? Create(IExcelSheet sheet)
        {
            return sheet switch
            {
                XlsxSheet xlsx => new XlsxPipelineSource(xlsx.Owner, xlsx.Index),
                XlsbSheet xlsb => new XlsbPipelineSource(xlsb.Owner, xlsb.Index),
                _ => null,
            };
        }

        internal static SheetPipelineSource? TryCreate(IExcelSheet sheet, int dop)
        {
            if (dop <= 1)
            {
                return null;
            }
            SheetPipelineSource? source = Create(sheet);
            return source is not null && source.UncompressedSize >= MinBytes ? source : null;
        }

        private protected static void Skip(Stream stream, long count)
        {
            byte[] scratch = ArrayPool<byte>.Shared.Rent(64 * 1024);
            try
            {
                while (count > 0)
                {
                    int read = stream.Read(scratch, 0, (int)Math.Min(count, scratch.Length));
                    if (read == 0)
                    {
                        return;
                    }
                    count -= read;
                }
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(scratch);
            }
        }
    }
}
