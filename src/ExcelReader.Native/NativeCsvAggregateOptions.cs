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
        internal static int Translate(NativeCsvParallelOptionsRaw? raw, out CsvParallelOptions options)
        {
            options = CsvParallelOptions.Default;
            if (raw is null)
            {
                return NativeStatus.Ok;
            }

            NativeCsvParallelOptionsRaw value = raw.Value;
            int expectedSize = sizeof(NativeCsvParallelOptionsRaw);
            if (value.StructSize != expectedSize
                || value.DegreeOfParallelism < 0
                || value.HeaderRow < 0
                || !NativeOptionDecode.TryByte(value.Delimiter, "csv_parallel_options", "delimiter", out byte? delimiter, out _)
                || !NativeOptionDecode.TryByte(value.Quote, "csv_parallel_options", "quote", out byte? quote, out _)
                || !NativeOptionDecode.TryState(value.DetectBom, "csv_parallel_options", "detect_bom", out bool? detectBom, out _)
                || value.MaxCellBytes < 0)
            {
                return NativeStatus.InvalidArgument;
            }

            CsvReaderOptions reader = CsvReaderOptions.Default;
            if (delimiter is byte d)
            {
                reader = reader with { Delimiter = d };
            }
            if (quote is byte q)
            {
                reader = reader with { Quote = q };
            }
            if (detectBom is bool bom)
            {
                reader = reader with { DetectEncodingFromByteOrderMark = bom };
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
