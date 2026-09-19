using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
        internal static int InferSchema(NativeHandle? handle, int headerRow, int sampleSize, out NativeInferredSchema schema)
        {
            schema = default;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }
            if (headerRow < 0)
            {
                SetLastError($"header_row must be 0 (no header) or a positive row number; got {headerRow}.");
                return NativeStatus.InvalidArgument;
            }
            if (sampleSize <= 0)
            {
                SetLastError($"sample_size must be positive; got {sampleSize}.");
                return NativeStatus.InvalidArgument;
            }

            ClearLastError();
            IExcelRowEnumerator? rows = null;
            try
            {
                handle.FaultLiveSession("xl_infer_schema");
                rows = handle.Reader.GetEnumerator();
                schema = BuildSchema(SchemaInference.Infer(rows, handle.Reader.IsDate1904, headerRow, sampleSize));
                return NativeStatus.Ok;
            }
            catch (ArgumentException exception)
            {
                SetLastError(exception.Message);
                schema = default;
                return NativeStatus.InvalidArgument;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                schema = default;
                return NativeStatus.Error;
            }
            finally
            {
                rows?.Dispose();
            }
        }

        internal static void FreeSchema(ref NativeInferredSchema schema)
        {
            if (schema.Columns == IntPtr.Zero)
            {
                schema = default;
                return;
            }

            NativeColumnSpecRaw* columns = (NativeColumnSpecRaw*)schema.Columns;
            for (int i = 0; i < schema.ColumnCount; i++)
            {
                NativeColumnSpecRaw spec = columns[i];
                if (spec.NameCount > 0 && spec.Names is not null)
                {
                    Marshal.FreeHGlobal((IntPtr)spec.Names[0]);
                    Marshal.FreeHGlobal((IntPtr)spec.Names);
                    Marshal.FreeHGlobal((IntPtr)spec.NameLens);
                }
            }
            Marshal.FreeHGlobal(schema.Columns);
            schema = default;
        }

        private static NativeInferredSchema BuildSchema(ExcelColumnSchema[] columns)
        {
            if (columns.Length == 0)
            {
                return new NativeInferredSchema { Columns = IntPtr.Zero, ColumnCount = 0 };
            }

            NativeColumnSpecRaw* block = (NativeColumnSpecRaw*)Marshal.AllocHGlobal(checked(columns.Length * sizeof(NativeColumnSpecRaw)));
            for (int i = 0; i < columns.Length; i++)
            {
                block[i] = BuildSpec(columns[i]);
            }
            return new NativeInferredSchema { Columns = (IntPtr)block, ColumnCount = columns.Length };
        }

        private static NativeColumnSpecRaw BuildSpec(ExcelColumnSchema column)
        {
            byte** namesBlock = null;
            int* lensBlock = null;
            int nameCount = 0;
            if (column.Name is not null)
            {
                int nameLen = Encoding.UTF8.GetByteCount(column.Name);
                byte* namePtr = (byte*)Marshal.AllocHGlobal(Math.Max(nameLen, 1));
                Encoding.UTF8.GetBytes(column.Name, new Span<byte>(namePtr, nameLen));

                namesBlock = (byte**)Marshal.AllocHGlobal(sizeof(byte*));
                namesBlock[0] = namePtr;
                lensBlock = (int*)Marshal.AllocHGlobal(sizeof(int));
                lensBlock[0] = nameLen;
                nameCount = 1;
            }
            return new NativeColumnSpecRaw
            {
                Names = namesBlock,
                NameLens = lensBlock,
                NameCount = nameCount,
                Index = column.Index,
                Type = (int)column.Type,
                Nullable = column.IsNullable ? 1 : 0,
            };
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeInferredSchema
    {
        public IntPtr Columns;
        public int ColumnCount;
    }
}
