using System.Buffers;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Parser.SheetPipeline
{
    internal readonly record struct SheetChunk(byte[] Buffer, int Length);

    internal static class ChunkProducer
    {
        internal static async Task RunAsync(
            Stream source, RowCutter cutter, int chunkBytes, int maxCellBytes, Func<SheetChunk, ValueTask> emit, CancellationToken ct)
        {
            byte[]? buffer = ArrayPool<byte>.Shared.Rent(chunkBytes);
            int target = chunkBytes;
            int filled = 0;
            try
            {
                while (true)
                {
                    bool eof = false;
                    while (filled < target)
                    {
                        int read = await source.ReadAsync(buffer.AsMemory(filled, target - filled), ct).ConfigureAwait(false);
                        if (read == 0)
                        {
                            eof = true;
                            break;
                        }
                        filled += read;
                    }

                    int cut = cutter.Cut(buffer.AsSpan(0, filled), out bool ended);
                    if (ended || eof)
                    {
                        int length = ended ? cut : filled;
                        if (length > 0)
                        {
                            byte[] last = buffer;
                            buffer = null;
                            await emit(new SheetChunk(last, length)).ConfigureAwait(false);
                        }
                        return;
                    }
                    if (cut < 0)
                    {
                        target = LimitChecks.NextBufferSize(maxCellBytes, nameof(ExcelReaderOptions.MaxCellBytes), target, target + 1);
                        if (buffer.Length < target)
                        {
                            byte[] bigger = ArrayPool<byte>.Shared.Rent(target);
                            buffer.AsSpan(0, filled).CopyTo(bigger);
                            ArrayPool<byte>.Shared.Return(buffer);
                            buffer = bigger;
                        }
                        continue;
                    }

                    int tail = filled - cut;
                    byte[] next = ArrayPool<byte>.Shared.Rent(Math.Max(chunkBytes, tail));
                    buffer.AsSpan(cut, tail).CopyTo(next);
                    byte[] full = buffer;
                    buffer = next;
                    filled = tail;
                    target = Math.Max(chunkBytes, tail);
                    cutter.Rebase(cut);
                    await emit(new SheetChunk(full, cut)).ConfigureAwait(false);
                }
            }
            finally
            {
                if (buffer is not null)
                {
                    ArrayPool<byte>.Shared.Return(buffer);
                }
            }
        }
    }
}
