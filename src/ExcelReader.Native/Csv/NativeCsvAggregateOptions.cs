using System.Runtime.InteropServices;
using ExcelReader.Core.Reader.Csv;

namespace ExcelReader.Native.Csv
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
        private const string OptionsName = "csv_parallel_options";

        internal static int Translate(NativeCsvParallelOptionsRaw? raw, out CsvParallelOptions options)
        {
            options = CsvParallelOptions.Default;
            if (raw is null)
            {
                return NativeStatus.Ok;
            }

            NativeCsvParallelOptionsRaw value = raw.Value;
            int expectedSize = sizeof(NativeCsvParallelOptionsRaw);
            if (value.StructSize != expectedSize)
            {
                NativeApi.SetLastError(
                    $"{OptionsName}.struct_size must be exactly {expectedSize}; got {value.StructSize}. Pass NULL for defaults, not a zeroed struct.");
                return NativeStatus.InvalidArgument;
            }
            if (value.DegreeOfParallelism < 0)
            {
                NativeApi.SetLastError($"{OptionsName}.degree_of_parallelism must be 0 (processor count) or positive; got {value.DegreeOfParallelism}.");
                return NativeStatus.InvalidArgument;
            }
            if (value.HeaderRow < 0)
            {
                NativeApi.SetLastError($"{OptionsName}.header_row must be 0 (no header) or a 1-based record number; got {value.HeaderRow}.");
                return NativeStatus.InvalidArgument;
            }
            if (value.MaxCellBytes < 0)
            {
                NativeApi.SetLastError($"{OptionsName}.max_cell_bytes must be 0 (default) or positive; got {value.MaxCellBytes}.");
                return NativeStatus.InvalidArgument;
            }
            if (!NativeOptionDecode.TryByte(value.Delimiter, OptionsName, "delimiter", out byte? delimiter, out string? error)
                || !NativeOptionDecode.TryByte(value.Quote, OptionsName, "quote", out byte? quote, out error)
                || !NativeOptionDecode.TryState(value.DetectBom, OptionsName, "detect_bom", out bool? detectBom, out error))
            {
                NativeApi.SetLastError(error!);
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
