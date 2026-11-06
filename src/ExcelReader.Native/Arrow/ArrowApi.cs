using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Reader.Columnar;
using ExcelReader.Core.Reader.Schema;
using ExcelReader.Native.Typed;

namespace ExcelReader.Native.Arrow
{
    internal static unsafe partial class ArrowApi
    {
        internal static int ParseArrow(NativeHandle? handle, int sheet, NativeColumnSpec[] specs, int headerRow, out ArrowArray array, out ArrowSchema schema)
        {
            return ParseArrow(handle, sheet, specs, headerRow, 1, out array, out schema);
        }

        internal static int ParseArrow(NativeHandle? handle, int sheet, NativeColumnSpec[] specs, int headerRow, int degreeOfParallelism,
            out ArrowArray array, out ArrowSchema schema)
        {
            array = default;
            schema = default;
            int status = TypedApi.ParseColumns(handle, sheet, specs, headerRow, degreeOfParallelism, allowDictionary: true, chunkSizeOverride: 0,
                out ColumnarTable? columns);
            if (status != NativeStatus.Ok)
            {
                return status;
            }

            using ColumnarTable table = columns!;
            try
            {
                schema = BuildArrowSchema(specs);
                array = BuildArrowArray(table);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                fixed (ArrowSchema* pinned = &schema)
                {
                    ReleaseArrowSchema((IntPtr)pinned);
                }
                array = default;
                schema = default;
                return NativeStatus.Error;
            }
        }

        internal static void ReleaseArrowSchema(IntPtr schemaPtr)
        {
            if (schemaPtr == IntPtr.Zero)
            {
                return;
            }
            ArrowSchema schema = Marshal.PtrToStructure<ArrowSchema>(schemaPtr);
            if (schema.Release == IntPtr.Zero)
            {
                return;
            }

            ReleaseChildren(schema.Children, schema.NChildren, &ReleaseArrowSchema);
            ReleaseDictionary(schema.Dictionary, &ReleaseArrowSchema);
            FreeIfSet(schema.Format);
            FreeIfSet(schema.Name);

            schema.Release = IntPtr.Zero;
            Marshal.StructureToPtr(schema, schemaPtr, false);
        }

        internal static void ReleaseArrowArray(IntPtr arrayPtr)
        {
            if (arrayPtr == IntPtr.Zero)
            {
                return;
            }
            ArrowArray array = Marshal.PtrToStructure<ArrowArray>(arrayPtr);
            if (array.Release == IntPtr.Zero)
            {
                return;
            }

            ReleaseChildren(array.Children, array.NChildren, &ReleaseArrowArray);
            ReleaseDictionary(array.Dictionary, &ReleaseArrowArray);
            for (long i = 0; i < array.NBuffers; i++)
            {
                FreeIfSet(Marshal.ReadIntPtr(array.Buffers, (int)(i * IntPtr.Size)));
            }
            if (array.NBuffers > 0)
            {
                Marshal.FreeHGlobal(array.Buffers);
            }

            array.Release = IntPtr.Zero;
            Marshal.StructureToPtr(array, arrayPtr, false);
        }

        private static void ReleaseChildren(IntPtr children, long count, delegate*<IntPtr, void> release)
        {
            for (long i = 0; i < count; i++)
            {
                IntPtr child = Marshal.ReadIntPtr(children, (int)(i * IntPtr.Size));
                release(child);
                Marshal.FreeHGlobal(child);
            }
            if (count > 0)
            {
                Marshal.FreeHGlobal(children);
            }
        }

        private static void ReleaseDictionary(IntPtr dictionary, delegate*<IntPtr, void> release)
        {
            if (dictionary == IntPtr.Zero)
            {
                return;
            }
            release(dictionary);
            Marshal.FreeHGlobal(dictionary);
        }

        private static void FreeIfSet(IntPtr pointer)
        {
            if (pointer != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(pointer);
            }
        }

        private static ArrowSchema BuildArrowSchema(NativeColumnSpec[] specs)
        {
            int count = specs.Length;
            IntPtr children = Marshal.AllocHGlobal(checked(count * IntPtr.Size));
            int built = 0;
            try
            {
                for (; built < count; built++)
                {
                    IntPtr child = Marshal.AllocHGlobal(Marshal.SizeOf<ArrowSchema>());
                    Marshal.StructureToPtr(BuildChildSchema(specs[built], built), child, false);
                    Marshal.WriteIntPtr(children, built * IntPtr.Size, child);
                }

                return new ArrowSchema
                {
                    Format = AllocUtf8Z("+s"),
                    NChildren = count,
                    Children = children,
                    Release = SchemaReleaseCallback,
                };
            }
            catch
            {
                for (int i = 0; i < built; i++)
                {
                    IntPtr child = Marshal.ReadIntPtr(children, i * IntPtr.Size);
                    ReleaseArrowSchema(child);
                    Marshal.FreeHGlobal(child);
                }
                Marshal.FreeHGlobal(children);
                throw;
            }
        }

        private static ArrowSchema BuildChildSchema(NativeColumnSpec spec, int index)
        {
            string name = spec.Names.Length > 0 ? spec.Names[0] : index.ToString(CultureInfo.InvariantCulture);
            ArrowSchema child = new()
            {
                Format = AllocUtf8Z(ArrowFormatCode(spec.Type)),
                Name = AllocUtf8Z(name),
                Flags = spec.Nullable ? ArrowFlags.Nullable : 0,
                Release = SchemaReleaseCallback,
            };
            if (spec.Type == NativeColumnType.StringDictionary)
            {
                child.Dictionary = AllocStruct(new ArrowSchema { Format = AllocUtf8Z("u"), Release = SchemaReleaseCallback });
            }
            return child;
        }

        private static IntPtr AllocStruct<T>(T value) where T : unmanaged
        {
            IntPtr block = Marshal.AllocHGlobal(sizeof(T));
            *(T*)block = value;
            return block;
        }

        private static ArrowArray BuildArrowArray(ColumnarTable table)
        {
            int count = table.Columns.Length;
            IntPtr children = Marshal.AllocHGlobal(checked(count * IntPtr.Size));
            int built = 0;
            try
            {
                for (; built < count; built++)
                {
                    IntPtr child = Marshal.AllocHGlobal(Marshal.SizeOf<ArrowArray>());
                    Marshal.StructureToPtr(BuildChildArray(ref table.Columns[built]), child, false);
                    Marshal.WriteIntPtr(children, built * IntPtr.Size, child);
                }

                IntPtr buffers = Marshal.AllocHGlobal(IntPtr.Size);
                Marshal.WriteIntPtr(buffers, 0, IntPtr.Zero);

                return new ArrowArray
                {
                    Length = table.RowCount,
                    NBuffers = 1,
                    Buffers = buffers,
                    NChildren = count,
                    Children = children,
                    Release = ArrayReleaseCallback,
                };
            }
            catch
            {
                for (int i = 0; i < built; i++)
                {
                    IntPtr child = Marshal.ReadIntPtr(children, i * IntPtr.Size);
                    ReleaseArrowArray(child);
                    Marshal.FreeHGlobal(child);
                }
                Marshal.FreeHGlobal(children);
                throw;
            }
        }

        private static ArrowArray BuildChildArray(ref BuiltColumn column)
        {
            if (column.Type == ExcelColumnType.StringDictionaryColumn)
            {
                return BuildDictionaryArray(ref column);
            }
            bool isString = column.Type == ExcelColumnType.StringColumn;
            int bufferCount = isString ? 3 : 2;
            IntPtr* buffers = (IntPtr*)Marshal.AllocHGlobal(bufferCount * IntPtr.Size);
            try
            {
                buffers[1] = ValuesBuffer(ref column);
            }
            catch
            {
                Marshal.FreeHGlobal((IntPtr)buffers);
                throw;
            }
            if (isString)
            {
                buffers[2] = TakeBlock(ref column.Data);
            }
            buffers[0] = TakeBlock(ref column.Validity);
            return Primitive(column.Length, buffers, bufferCount);
        }

        // Arrow packs booleans to bits; every other type hands its block over as it is.
        private static IntPtr ValuesBuffer(ref BuiltColumn column)
        {
            if (column.Type == ExcelColumnType.BoolColumn)
            {
                return BitPackBoolColumn(column.Values, column.Length);
            }
            return TakeBlock(ref column.Values);
        }

        // Indices with the column's validity; the distinct strings hang off them as a utf8 array with no nulls.
        private static ArrowArray BuildDictionaryArray(ref BuiltColumn column)
        {
            IntPtr* values = (IntPtr*)Marshal.AllocHGlobal(3 * IntPtr.Size);
            values[0] = IntPtr.Zero;
            values[1] = TakeBlock(ref column.DictionaryOffsets);
            values[2] = TakeBlock(ref column.Data);
            ArrowArray dictionary = Primitive(column.DictionaryLength, values, 3);
            IntPtr* indices = (IntPtr*)Marshal.AllocHGlobal(2 * IntPtr.Size);
            indices[1] = TakeBlock(ref column.Values);
            indices[0] = TakeBlock(ref column.Validity);
            ArrowArray array = Primitive(column.Length, indices, 2);
            array.Dictionary = AllocStruct(dictionary);
            return array;
        }

        private static ArrowArray Primitive(long length, IntPtr* buffers, int bufferCount)
        {
            return new ArrowArray
            {
                Length = length,
                NullCount = ColumnBlocks.CountUnset(buffers[0], length),
                NBuffers = bufferCount,
                Buffers = (IntPtr)buffers,
                Release = ArrayReleaseCallback,
            };
        }

        private static IntPtr TakeBlock(ref IntPtr block)
        {
            IntPtr taken = block;
            block = IntPtr.Zero;
            return taken;
        }

        private static IntPtr BitPackBoolColumn(IntPtr byteValues, long length)
        {
            return ColumnBlocks.PackBitsLsbFirst(new ReadOnlySpan<byte>((void*)byteValues, (int)length));
        }

        private static string ArrowFormatCode(int type)
        {
            return type switch
            {
                NativeColumnType.String => "u",
                NativeColumnType.StringDictionary => "i",
                NativeColumnType.Int64 => "l",
                NativeColumnType.Float64 => "g",
                NativeColumnType.Bool => "b",
                NativeColumnType.Date => "tdD",
                NativeColumnType.Time => "ttu",
                _ => "tsu:",
            };
        }

        private static IntPtr AllocUtf8Z(string value)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(value);
            IntPtr block = Marshal.AllocHGlobal(bytes.Length + 1);
            Marshal.Copy(bytes, 0, block, bytes.Length);
            Marshal.WriteByte(block, bytes.Length, 0);
            return block;
        }

        private static IntPtr SchemaReleaseCallback
        {
            get
            {
                return (IntPtr)(delegate* unmanaged<ArrowSchema*, void>)&Exports.ReleaseArrowSchemaCallback;
            }
        }

        private static IntPtr ArrayReleaseCallback
        {
            get
            {
                return (IntPtr)(delegate* unmanaged<ArrowArray*, void>)&Exports.ReleaseArrowArrayCallback;
            }
        }
    }
}
