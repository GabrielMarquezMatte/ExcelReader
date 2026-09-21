using System.Globalization;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
        internal static int ParseArrow(NativeHandle? handle, NativeColumnSpec[] specs, int headerRow, out ArrowArray array, out ArrowSchema schema)
        {
            array = default;
            schema = default;
            int status = TypedParseSession.OpenTransient(handle, specs, headerRow, "xl_parse_arrow",
                out TypedParseSession? session);
            if (status != NativeStatus.Ok)
            {
                return status;
            }

            NativeTable table = default;
            try
            {
                using (TypedParseSession open = session!)
                {
                    status = open.NextBatch(out table);
                    if (status == NativeStatus.Eof)
                    {
                        table = BuildEmptyTable(specs);
                        status = NativeStatus.Ok;
                    }
                }
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
            if (status != NativeStatus.Ok)
            {
                return status;
            }

            try
            {
                schema = BuildArrowSchema(specs);
                array = BuildArrowArray(specs, table);
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                fixed (ArrowSchema* pinned = &schema)
                {
                    ReleaseArrowSchema((IntPtr)pinned);
                }
                array = default;
                schema = default;
                return NativeStatus.Error;
            }
            finally
            {
                FreeTable(ref table);
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
            return new ArrowSchema
            {
                Format = AllocUtf8Z(ArrowFormatCode(spec.Type)),
                Name = AllocUtf8Z(name),
                Flags = spec.Nullable ? ArrowFlags.Nullable : 0,
                Release = SchemaReleaseCallback,
            };
        }

        private static ArrowArray BuildArrowArray(NativeColumnSpec[] specs, NativeTable table)
        {
            int count = table.ColumnCount;
            IntPtr children = Marshal.AllocHGlobal(checked(count * IntPtr.Size));
            int built = 0;
            try
            {
                for (; built < count; built++)
                {
                    IntPtr child = Marshal.AllocHGlobal(Marshal.SizeOf<ArrowArray>());
                    Marshal.StructureToPtr(BuildChildArray(specs[built].Type, &((NativeColumn*)table.Columns)[built]), child, false);
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

        // Takes ownership of the column's buffers and clears them, so FreeTable leaves them to the Arrow release.
        private static ArrowArray BuildChildArray(int type, NativeColumn* column)
        {
            int bufferCount = type == NativeColumnType.String ? 3 : 2;
            IntPtr* buffers = (IntPtr*)Marshal.AllocHGlobal(bufferCount * IntPtr.Size);
            if (type == NativeColumnType.Bool)
            {
                try
                {
                    buffers[1] = BitPackBoolColumn(column->Values, column->Length);
                }
                catch
                {
                    Marshal.FreeHGlobal((IntPtr)buffers);
                    throw;
                }
            }
            else
            {
                buffers[1] = column->Values;
                column->Values = IntPtr.Zero;
            }
            if (type == NativeColumnType.String)
            {
                buffers[2] = column->Data;
                column->Data = IntPtr.Zero;
            }
            buffers[0] = column->Validity;
            column->Validity = IntPtr.Zero;

            return new ArrowArray
            {
                Length = column->Length,
                NullCount = CountUnset(buffers[0], column->Length),
                NBuffers = bufferCount,
                Buffers = (IntPtr)buffers,
                Release = ArrayReleaseCallback,
            };
        }

        private static long CountUnset(IntPtr validity, long length)
        {
            if (validity == IntPtr.Zero)
            {
                return 0;
            }
            ReadOnlySpan<byte> bits = new((void*)validity, (int)((length + 7) / 8));
            var ulongs = MemoryMarshal.Cast<byte, ulong>(bits);
            long set = 0;
            foreach (ref readonly ulong packed in ulongs)
            {
                set += BitOperations.PopCount(packed);
            }
            foreach (ref readonly byte packed in bits.Slice(ulongs.Length * sizeof(ulong)))
            {
                set += BitOperations.PopCount(packed);
            }
            return length - set;
        }

        private static IntPtr BitPackBoolColumn(IntPtr byteValues, long length)
        {
            return PackBitsLsbFirst(new ReadOnlySpan<byte>((void*)byteValues, (int)length));
        }

        private static string ArrowFormatCode(int type)
        {
            return type switch
            {
                NativeColumnType.String => "u",
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

        private static NativeColumn ColumnAt(NativeTable table, int index)
        {
            return ((NativeColumn*)table.Columns)[index];
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
