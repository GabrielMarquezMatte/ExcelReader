using System.IO.Compression;

namespace ExcelReader.Core.Writer.Parallel
{
    /// <summary>
    /// What a sheet writer offers the parallel record path. The first five members act on the real
    /// sheet, from the calling flow. The rest act on a detached copy, which formats batches into its
    /// own buffer on a pool thread and never touches the output.
    /// </summary>
    internal interface IBatchSheet<TSelf, TRow>
        where TSelf : class, IBatchSheet<TSelf, TRow>
        where TRow : IRowWriter
    {
        CompressionLevel BatchCompression { get; }

        /// <summary>Starts a parallel stretch and returns how many rows the sheet already holds.</summary>
        ValueTask<int> BeginSegmentsAsync(CancellationToken ct);

        /// <summary>Writes one batch to the sheet's entry; called in batch order.</summary>
        ValueTask AppendSegmentAsync(BatchSegment segment, CancellationToken ct);

        void EndSegments();

        /// <summary>Faults the sheet after a failed parallel stretch. Never throws.</summary>
        ValueTask FailSegmentsAsync();

        TSelf CreateDetached();

        void BeginBatch(int rowsBefore);

        TRow StartBatchRow();

        /// <summary>Ends the batch, records its used range in <paramref name="segment"/> and returns its bytes.</summary>
        ReadOnlySpan<byte> EndBatch(BatchSegment segment);

        void ReleaseDetached();
    }
}
