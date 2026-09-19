using System.Buffers;
using System.Text;

namespace ExcelReader.Core.Reader
{
    internal static class XlsbSharedStrings
    {
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
