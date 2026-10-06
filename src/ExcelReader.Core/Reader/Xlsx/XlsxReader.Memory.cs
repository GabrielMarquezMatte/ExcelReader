using System.Buffers;
using ExcelReader.Core.Reader.Internal;
using ExcelReader.Core.Reader.Zip;

namespace ExcelReader.Core.Reader.Xlsx
{
    public sealed partial class XlsxReader
    {
        private void ParseSharedFromMemory(ReadOnlyMemory<byte> content, long entryLength)
        {
            LimitChecks.ThrowIfEntryLengthExceeds(entryLength, Array.MaxLength, "ArrayMaxLength");
            int partLength = (int)entryLength;
            int growthCap = SharedFlatGrowthCap();
            var io = new BufferedStreamCursor(content, growthCap, nameof(ExcelReaderOptions.MaxSharedStringBytes));
            _sharedFlat = ArrayPool<byte>.Shared.Rent(Math.Max(1, partLength));
            _sharedOffsets = ParseSharedBody(io, stream: null, partLength);
        }
    }
}
