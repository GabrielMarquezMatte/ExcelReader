using System.Buffers;
using System.Text;

namespace ExcelReader.Core.Reader
{
    // Decodes xl/sharedStrings.bin (binary BIFF12) into a flat UTF-8 buffer + offsets, matching how
    // XlsxReader stores shared strings so Cell consumes UTF-8 either way:
    // string i = Flat[Offsets[i]..Offsets[i+1]].
    internal static class XlsbSharedStrings
    {
        // Used by the memory-workbook path (the whole part is already decompressed and in hand, so
        // there is no stream to read incrementally) and as the reference decoder XlsbPartsTests
        // checks the streaming overloads' output against.
        //
        // Builds through the same ArrayPool-backed growth ParseCore uses for the streamed variant
        // (AppendItemPooled/PrepareOffsets/AddOffset below), rather than a plain array trimmed with
        // Array.Resize at the end. UTF-8 almost always comes out smaller than the BIFF12 binary it
        // was decoded from, so that trim fired on essentially every call — two full-size GC
        // allocations per open (one of them immediately discarded), both large enough to land on the
        // LOH for any real workbook. Now only the one the caller actually keeps is GC-owned; the
        // growth buffer is a pool rental, returned once the exact-size result has been copied out.
        internal static (byte[] Flat, int[] Offsets) Parse(ReadOnlySpan<byte> sharedBin, ExcelReaderOptions? options = null)
        {
            ExcelReaderOptions effectiveOptions = options ?? ExcelReaderOptions.Default;
            if (sharedBin.IsEmpty)
            {
                return ([], [0]);
            }
            LimitChecks.ThrowIfOverSharedStringLimit(effectiveOptions, sharedBin.Length);

            int[] offsets = [0];
            int offsetCount = 1;
            byte[] flat = ArrayPool<byte>.Shared.Rent(Math.Max(16, sharedBin.Length));
            try
            {
                int flatLen = 0;
                var reader = new Biff12RecordReader(sharedBin);
                while (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
                {
                    if (id == Brt.BeginSst)
                    {
                        // The whole part is in hand, so its length stands in for the streaming path's
                        // entryLength — both bound PrepareOffsets' sanity check the same way.
                        PrepareOffsets(payload, sharedBin.Length, ref offsets, ref offsetCount);
                    }
                    else if (id == Brt.SSTItem)
                    {
                        flatLen = AppendItemPooled(payload, ref flat, flatLen, effectiveOptions);
                        AddOffset(ref offsets, ref offsetCount, flatLen);
                    }
                }
                byte[] result = new byte[flatLen];
                flat.AsSpan(0, flatLen).CopyTo(result);
                return (result, offsetCount == offsets.Length ? offsets : offsets[..offsetCount]);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(flat);
            }
        }

        // BrtSSTItem records carry their own payload length, so a BufferedStreamCursor can compact
        // everything before the current record while growing only for one unusually large record.
        // The returned flat buffer is rented and belongs to XlsbReader for its lifetime.
        internal static (byte[] Flat, int[] Offsets) ParseStreaming(Stream stream, long entryLength, ExcelReaderOptions options)
        {
            var io = new BufferedStreamCursor(GrowthCap(options), nameof(ExcelReaderOptions.MaxSharedStringBytes),
                WorkbookLookups.InitialBufferCapacity(entryLength));
            byte[] flat = ArrayPool<byte>.Shared.Rent(256);
            try
            {
                var result = ParseCore(io, stream, entryLength, options, ref flat);
                io.Return();
                return (flat, result);
            }
            catch
            {
                ArrayPool<byte>.Shared.Return(flat);
                io.Return();
                throw;
            }
        }

        internal static async ValueTask<(byte[] Flat, int[] Offsets)> ParseStreamingAsync(
            Stream stream, long entryLength, ExcelReaderOptions options, CancellationToken ct)
        {
            var io = new BufferedStreamCursor(GrowthCap(options), nameof(ExcelReaderOptions.MaxSharedStringBytes),
                WorkbookLookups.InitialBufferCapacity(entryLength));
            byte[] flat = ArrayPool<byte>.Shared.Rent(256);
            try
            {
                return await ParseCoreAsync(io, stream, entryLength, options, flat, ct).ConfigureAwait(false);
            }
            finally
            {
                io.Return();
            }
        }

        private static int[] ParseCore(BufferedStreamCursor io, Stream stream, long entryLength, ExcelReaderOptions options, ref byte[] flat)
        {
            int[] offsets = [0];
            int offsetCount = 1;
            int flatLen = 0;
            while (true)
            {
                var reader = new Biff12RecordReader(io.Buf.AsSpan(io.Pos, io.Len - io.Pos));
                if (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
                {
                    io.Pos += reader.Position;
                    if (id == Brt.BeginSst)
                    {
                        PrepareOffsets(payload, entryLength, ref offsets, ref offsetCount);
                    }
                    else if (id == Brt.SSTItem)
                    {
                        flatLen = AppendItemPooled(payload, ref flat, flatLen, options);
                        AddOffset(ref offsets, ref offsetCount, flatLen);
                    }
                    continue;
                }
                if (io.Eof)
                {
                    break;
                }
                io.Fill(stream);
            }
            return offsetCount == offsets.Length ? offsets : offsets[..offsetCount];
        }

        private static async ValueTask<(byte[] Flat, int[] Offsets)> ParseCoreAsync(
            BufferedStreamCursor io, Stream stream, long entryLength, ExcelReaderOptions options,
            byte[] flat, CancellationToken ct)
        {
            try
            {
                int[] offsets = [0];
                int offsetCount = 1;
                int flatLen = 0;
                while (true)
                {
                    var reader = new Biff12RecordReader(io.Buf.AsSpan(io.Pos, io.Len - io.Pos));
                    if (reader.TryReadRecord(out int id, out ReadOnlySpan<byte> payload))
                    {
                        io.Pos += reader.Position;
                        if (id == Brt.BeginSst)
                        {
                            PrepareOffsets(payload, entryLength, ref offsets, ref offsetCount);
                        }
                        else if (id == Brt.SSTItem)
                        {
                            flatLen = AppendItemPooled(payload, ref flat, flatLen, options);
                            AddOffset(ref offsets, ref offsetCount, flatLen);
                        }
                        continue;
                    }
                    if (io.Eof)
                    {
                        break;
                    }
                    await io.FillAsync(stream, ct).ConfigureAwait(false);
                }
                return (flat, offsetCount == offsets.Length ? offsets : offsets[..offsetCount]);
            }
            catch
            {
                ArrayPool<byte>.Shared.Return(flat);
                throw;
            }
        }

        // BeginSst stores total strings then unique strings. The latter lets the common well-formed
        // path allocate its offsets table once; do not trust it when it exceeds the physical maximum
        // number of seven-byte empty BrtSSTItem records the entry could contain.
        private static void PrepareOffsets(ReadOnlySpan<byte> payload, long entryLength, ref int[] offsets, ref int offsetCount)
        {
            if (payload.Length < 8)
            {
                return;
            }
            uint unique = Biff12.ReadU32(payload, 4);
            long maxItems = entryLength > 0 ? entryLength / 7 : 0;
            if (unique == 0 || unique > maxItems || unique >= Array.MaxLength)
            {
                return;
            }
            offsets = new int[(int)unique + 1];
            offsets[0] = 0;
            offsetCount = 1;
        }

        private static int AppendItemPooled(ReadOnlySpan<byte> payload, ref byte[] flat, int flatLen, ExcelReaderOptions options)
        {
            if (payload.Length < 1 || !Biff12.TryReadWideString(payload, 1, out ReadOnlySpan<char> chars, out _))
            {
                return flatLen;
            }
            int needed = checked(flatLen + Encoding.UTF8.GetByteCount(chars));
            LimitChecks.ThrowIfOverSharedStringLimit(options, needed);
            if (needed > flat.Length)
            {
                byte[] bigger = ArrayPool<byte>.Shared.Rent(LimitChecks.NextBufferSize(
                    GrowthCap(options), nameof(ExcelReaderOptions.MaxSharedStringBytes), flat.Length, needed));
                flat.AsSpan(0, flatLen).CopyTo(bigger);
                ArrayPool<byte>.Shared.Return(flat);
                flat = bigger;
            }
            return flatLen + Encoding.UTF8.GetBytes(chars, flat.AsSpan(flatLen));
        }

        private static void AddOffset(ref int[] offsets, ref int count, int value)
        {
            if (count == offsets.Length)
            {
                Array.Resize(ref offsets, offsets.Length * 2);
            }
            offsets[count++] = value;
        }

        private static int GrowthCap(ExcelReaderOptions options)
        {
            return options.MaxSharedStringBytes <= 0
                ? 0
                : (int)Math.Min(options.MaxSharedStringBytes, Array.MaxLength);
        }
    }
}
