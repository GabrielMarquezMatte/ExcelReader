using System.Runtime.InteropServices;
using ExcelReader.Core.Reader;

namespace ExcelReader.Native
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeCsvAggregationRaw
    {
        public int StructSize;
        public IntPtr Seed;
        public IntPtr Accumulate;
        public IntPtr Combine;
        public IntPtr FreeState;
        public IntPtr UserData;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeCsvParallelOptionsRaw
    {
        public int StructSize;
        public int DegreeOfParallelism;
        public int HeaderRow;
        public int Delimiter;
        public int Quote;
        public int DetectBom;
        public int MaxCellBytes;
    }

    internal static unsafe class NativeCsvAggregateOptions
    {
        private const int OptTrue = 2;

        internal static int Translate(NativeCsvParallelOptionsRaw? raw, out CsvParallelOptions options)
        {
            options = CsvParallelOptions.Default;
            if (raw is null)
            {
                return NativeStatus.Ok;
            }

            NativeCsvParallelOptionsRaw value = raw.Value;
            if (value.StructSize < sizeof(NativeCsvParallelOptionsRaw)
                || value.DegreeOfParallelism < 0
                || value.HeaderRow < 0
                || value.Delimiter is < 0 or > 255
                || value.Quote is < 0 or > 255
                || value.DetectBom is < 0 or > OptTrue
                || value.MaxCellBytes < 0)
            {
                return NativeStatus.InvalidArgument;
            }

            CsvReaderOptions reader = CsvReaderOptions.Default;
            if (value.Delimiter != 0)
            {
                reader = reader with { Delimiter = (byte)value.Delimiter };
            }
            if (value.Quote != 0)
            {
                reader = reader with { Quote = (byte)value.Quote };
            }
            if (value.DetectBom != 0)
            {
                reader = reader with { DetectEncodingFromByteOrderMark = value.DetectBom == OptTrue };
            }
            if (value.MaxCellBytes != 0)
            {
                reader = reader with { MaxCellBytes = value.MaxCellBytes };
            }

            options = new CsvParallelOptions
            {
                DegreeOfParallelism = value.DegreeOfParallelism,
                HeaderRow = value.HeaderRow,
                Reader = reader,
            };
            return NativeStatus.Ok;
        }
    }
}
