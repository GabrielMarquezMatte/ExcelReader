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
        /// <summary>
        /// Guesses a <see cref="ParseTyped"/>/<c>xl_parse_arrow</c> schema by sampling the WHOLE current
        /// sheet, from its first row — independent of, and never disturbing, the incremental cursor
        /// <see cref="NextRow"/>/<see cref="NextRowDecoded"/>/<see cref="ReadAllBlob"/> share on
        /// <paramref name="handle"/>. Every guess comes from the sampled cells' own <see cref="CellType"/>
        /// tag (the same one <see cref="ParseTyped"/> already trusts to convert values) — no text
        /// sniffing, and no new parsing logic beyond <see cref="ExcelCellReaders.Parsable{TValue}"/>,
        /// reused here only to tell an integral column from a fractional one.
        /// </summary>
        /// <param name="headerRow">Same meaning as in <see cref="ParseTyped"/>: 1-based row number to
        /// take column names from; 0 means "no header", so every returned spec is index-based.</param>
        /// <param name="sampleSize">How many rows after the header to inspect. Must be positive.</param>
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
                rows = handle.Reader.GetEnumerator();
                schema = BuildSchema(SchemaInference.Infer(rows, handle.Reader.IsDate1904, headerRow, sampleSize));
                return NativeStatus.Ok;
            }
            catch (ArgumentException exception)
            {
                // Core throws for an unreachable header row; the ABI reports it as a status code.
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

        /// <summary>Releases a result returned by <see cref="InferSchema"/> and resets it to zero. Safe on a zeroed value.</summary>
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
                // NameCount > 0 is supposed to imply Names/NameLens are non-null (BuildSpec always
                // allocates both together), but a caller that null-checked and swapped in its own
                // freed-and-nulled Names field between InferSchema and this call would otherwise
                // dereference a null byte**. FreeRows/FreeTable already check the array pointer
                // first for the same reason; this matches them.
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

        // Every allocation this function makes is handed to the caller inside the returned schema and
        // freed only by FreeSchema — a thrown exception between an AllocHGlobal and the assignment to
        // `schema` below would leak it, but nothing after the loop's own allocations can throw.
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
                // ExcelColumnType's underlying values ARE the XL_T_* constants, by design — see the
                // remarks on the enum. This cast is the whole translation.
                Type = (int)column.Type,
                Nullable = column.IsNullable ? 1 : 0,
            };
        }
    }

    /// <summary>Flat C ABI representation of the whole result of <see cref="NativeApi.InferSchema"/>.
    /// <see cref="Columns"/> is one allocation of <see cref="ColumnCount"/> <see cref="NativeColumnSpecRaw"/>
    /// values; each spec's own non-null <see cref="NativeColumnSpecRaw.Name"/> is a separate allocation,
    /// freed individually by <see cref="NativeApi.FreeSchema"/>.</summary>
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeInferredSchema
    {
        public IntPtr Columns;
        public int ColumnCount;
    }
}
