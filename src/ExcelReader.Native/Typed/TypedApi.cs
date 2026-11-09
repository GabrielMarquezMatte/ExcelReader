using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Columnar;
using ExcelReader.Core.Reader.Schema;
using ExcelReader.Native.Reading;

namespace ExcelReader.Native.Typed
{
    internal static unsafe partial class TypedApi
    {
        [ThreadStatic]
        internal static bool LastParseRanInParallel;

        internal static int ParseTyped(NativeHandle? handle, int sheet, NativeColumnSpec[] specs, int headerRow, out NativeTable table)
        {
            return ParseTypedTable(handle, sheet, specs, headerRow, 1, out table);
        }

        internal static int ParseTypedTable(NativeHandle? handle, int sheet, NativeColumnSpec[] specs, int headerRow, int degreeOfParallelism,
            out NativeTable table, int chunkSizeOverride = 0)
        {
            table = default;
            int status = ParseColumns(handle, sheet, specs, headerRow, degreeOfParallelism, allowDictionary: false, chunkSizeOverride, out ColumnarTable? columns);
            if (status != NativeStatus.Ok)
            {
                return status;
            }
            using ColumnarTable built = columns!;
            try
            {
                table = ToNativeTable(built);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int ParseColumns(NativeHandle? handle, int sheet, NativeColumnSpec[] specs, int headerRow, int degreeOfParallelism,
            bool allowDictionary, int chunkSizeOverride, out ColumnarTable? columns)
        {
            columns = null;
            LastParseRanInParallel = false;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }
            if (degreeOfParallelism < 0)
            {
                NativeApi.SetLastError($"degree_of_parallelism must be 0 (processor count) or positive; got {degreeOfParallelism}.");
                return NativeStatus.InvalidArgument;
            }
            int status = ValidateSpecs(specs, headerRow, allowDictionary, out TypedColumnSpec[] coreSpecs);
            if (status != NativeStatus.Ok)
            {
                return status;
            }
            status = ReadApi.ResolveSheet(handle, sheet, out IExcelSheet? resolved);
            if (status != NativeStatus.Ok)
            {
                return status;
            }
            try
            {
                columns = TypedTableParser.Parse(resolved!, coreSpecs, headerRow, degreeOfParallelism, allowDictionary, chunkSizeOverride,
                    out LastParseRanInParallel, out string? specError);
                return columns is null ? InvalidArgument(specError!) : NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }

        internal static int ValidateSpecs(NativeColumnSpec[] specs, int headerRow, bool allowDictionary, out TypedColumnSpec[] coreSpecs)
        {
            coreSpecs = Array.ConvertAll(specs, ToCore);
            if (specs.Length == 0)
            {
                return InvalidArgument("xl_parse_typed requires at least one column spec.");
            }
            if (!TypedTableParser.TryValidate(coreSpecs, headerRow, allowDictionary, out string? error))
            {
                return InvalidArgument(error);
            }
            NativeApi.ClearLastError();
            return NativeStatus.Ok;
        }

        private static int InvalidArgument(string message)
        {
            NativeApi.SetLastError(message);
            return NativeStatus.InvalidArgument;
        }

        private static TypedColumnSpec ToCore(NativeColumnSpec spec)
        {
            return new TypedColumnSpec { Names = spec.Names, Index = spec.Index, Type = (ExcelColumnType)spec.Type, Nullable = spec.Nullable };
        }

        internal static NativeTable ToNativeTable(ColumnarTable columns)
        {
            int count = columns.Columns.Length;
            IntPtr block = ColumnBlocks.Alloc(checked(Math.Max(count, 1) * sizeof(NativeColumn)));
            NativeColumn* target = (NativeColumn*)block;
            for (int c = 0; c < count; c++)
            {
                ref BuiltColumn source = ref columns.Columns[c];
                target[c] = new NativeColumn
                {
                    Type = (int)source.Type,
                    Length = source.Length,
                    Values = source.Values,
                    Validity = source.Validity,
                    Data = source.Data,
                    DataLen = source.DataLen,
                };
                source = default;
            }
            return new NativeTable { ColumnCount = count, RowCount = columns.RowCount, Columns = block };
        }

        internal static void FreeTable(ref NativeTable table)
        {
            if (table.Columns == IntPtr.Zero)
            {
                table = default;
                return;
            }

            for (int index = 0; index < table.ColumnCount; index++)
            {
                ref readonly NativeColumn column = ref table.ColumnAt(index);
                ColumnBlocks.Free(column.Values);
                ColumnBlocks.Free(column.Validity);
                ColumnBlocks.Free(column.Data);
            }
            ColumnBlocks.Free(table.Columns);
            table = default;
        }

        internal static bool IsValidSpecCount(int specCount)
        {
            return specCount is > 0 and <= NativeLimits.MaxColumnSpecs;
        }

        internal static bool IsValidNameLength(int nameLength)
        {
            return nameLength is >= 0 and <= NativeLimits.MaxColumnNameBytes;
        }

        internal static bool IsValidNameCount(int nameCount)
        {
            return nameCount is >= 0 and <= NativeLimits.MaxNamesPerSpec;
        }
    }
}
