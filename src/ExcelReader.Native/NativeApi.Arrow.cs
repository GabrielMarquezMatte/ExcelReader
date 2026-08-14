using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelReader.Native
{
    internal static unsafe partial class NativeApi
    {
        /// <summary>
        /// Same schema-driven read as <see cref="ParseTyped"/>, exported as one top-level Arrow struct
        /// array/schema instead of a <see cref="NativeTable"/> — see excelreader_arrow.h for the
        /// XL_T_*-to-Arrow-format-code mapping and the ownership contract (the caller releases via
        /// <c>out_array-&gt;release</c>/<c>out_schema-&gt;release</c>, never <see cref="FreeTable"/>).
        /// </summary>
        /// <remarks>
        /// Not zero-copy from <see cref="ParseTyped"/>'s own buffers: every Arrow buffer here is a
        /// fresh allocation copied from the intermediate <see cref="NativeTable"/>, which is freed
        /// immediately after (including the deliberate <see cref="NativeColumnType.Bool"/> repack from
        /// one byte per row to Arrow's bit-packed layout). This trades one extra copy of already-computed
        /// columnar data for an ownership story with no buffers shared between two different release
        /// paths — true zero-copy would require building Arrow-shaped buffers during the row loop
        /// itself, which is a larger change deferred until zero-copy is actually measured to matter.
        /// </remarks>
        internal static int ParseArrow(NativeHandle? handle, NativeColumnSpec[] specs, int headerRow, out ArrowArray array, out ArrowSchema schema)
        {
            array = default;
            schema = default;
            int status = ParseTyped(handle, specs, headerRow, out NativeTable table);
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
                // schema may already hold a partially- or fully-built result if BuildArrowArray threw
                // after BuildArrowSchema succeeded — release it via the same path a real consumer would
                // use. ReleaseArrowSchema takes an address, not a value, so round-trip through a
                // throwaway native block rather than pinning this local.
                IntPtr temp = Marshal.AllocHGlobal(Marshal.SizeOf<ArrowSchema>());
                try
                {
                    Marshal.StructureToPtr(schema, temp, false);
                    ReleaseArrowSchema(temp);
                }
                finally
                {
                    Marshal.FreeHGlobal(temp);
                }
                array = default;
                schema = default;
                return NativeStatus.Error;
            }
            finally
            {
                // Arrow now owns independent copies of everything it needs (or, on failure, owns
                // nothing) — the intermediate table is never referenced by the caller either way.
                FreeTable(ref table);
            }
        }

        /// <summary>Releases a result returned by <see cref="ParseArrow"/>'s <paramref name="schemaPtr"/>
        /// half. Safe to call on an already-released (zeroed <see cref="ArrowSchema.Release"/>) schema —
        /// mirrors every other <c>xl_free_*</c>'s idempotency, required here because Arrow's own contract
        /// permits (and some consumers do) a defensive double-release check.</summary>
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

            for (long i = 0; i < schema.NChildren; i++)
            {
                IntPtr child = Marshal.ReadIntPtr(schema.Children, (int)(i * IntPtr.Size));
                ReleaseArrowSchema(child);
                Marshal.FreeHGlobal(child);
            }
            if (schema.NChildren > 0)
            {
                Marshal.FreeHGlobal(schema.Children);
            }
            FreeIfSet(schema.Format);
            FreeIfSet(schema.Name);

            schema.Release = IntPtr.Zero;
            Marshal.StructureToPtr(schema, schemaPtr, false);
        }

        /// <summary>The <see cref="ArrowArray"/> half of <see cref="ReleaseArrowSchema"/> — same
        /// recursive-children, then-buffers, then-idempotent-mark shape.</summary>
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

            for (long i = 0; i < array.NChildren; i++)
            {
                IntPtr child = Marshal.ReadIntPtr(array.Children, (int)(i * IntPtr.Size));
                ReleaseArrowArray(child);
                Marshal.FreeHGlobal(child);
            }
            if (array.NChildren > 0)
            {
                Marshal.FreeHGlobal(array.Children);
            }
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
            for (int i = 0; i < count; i++)
            {
                IntPtr child = Marshal.AllocHGlobal(Marshal.SizeOf<ArrowSchema>());
                Marshal.StructureToPtr(BuildChildSchema(specs[i], i), child, false);
                Marshal.WriteIntPtr(children, i * IntPtr.Size, child);
            }

            return new ArrowSchema
            {
                Format = AllocUtf8Z("+s"),
                NChildren = count,
                Children = children,
                Release = SchemaReleaseCallback,
            };
        }

        private static ArrowSchema BuildChildSchema(NativeColumnSpec spec, int index)
        {
            string name = spec.Name ?? index.ToString(CultureInfo.InvariantCulture);
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
            for (int i = 0; i < count; i++)
            {
                NativeColumn column = ColumnAt(table, i);
                IntPtr child = Marshal.AllocHGlobal(Marshal.SizeOf<ArrowArray>());
                Marshal.StructureToPtr(BuildChildArray(specs[i].Type, column), child, false);
                Marshal.WriteIntPtr(children, i * IntPtr.Size, child);
            }

            // The top-level struct array has exactly one buffer slot (validity), per the Arrow
            // convention for struct-typed arrays; a table-of-columns has no row-level nulls of its
            // own, so that slot is always the zero/absent pointer.
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

        private static ArrowArray BuildChildArray(int type, NativeColumn column)
        {
            IntPtr validity = CopyOptionalBuffer(column.Validity, (column.Length + 7) / 8);
            IntPtr[] dataBuffers = type switch
            {
                NativeColumnType.String => [CopyBuffer(column.Values, (column.Length + 1) * sizeof(int)), CopyBuffer(column.Data, column.DataLen)],
                NativeColumnType.Bool => [BitPackBoolColumn(column.Values, column.Length)],
                NativeColumnType.Float64 => [CopyBuffer(column.Values, column.Length * sizeof(double))],
                NativeColumnType.Date => [CopyBuffer(column.Values, column.Length * sizeof(int))],
                _ => [CopyBuffer(column.Values, column.Length * sizeof(long))], // Int64, Time, Timestamp
            };

            int bufferCount = 1 + dataBuffers.Length;
            IntPtr buffers = Marshal.AllocHGlobal(checked(bufferCount * IntPtr.Size));
            Marshal.WriteIntPtr(buffers, 0, validity);
            for (int i = 0; i < dataBuffers.Length; i++)
            {
                Marshal.WriteIntPtr(buffers, (i + 1) * IntPtr.Size, dataBuffers[i]);
            }

            return new ArrowArray
            {
                Length = column.Length,
                NullCount = CountUnset(column.Validity, column.Length),
                NBuffers = bufferCount,
                Buffers = buffers,
                Release = ArrayReleaseCallback,
            };
        }

        private static long CountUnset(IntPtr validity, long length)
        {
            if (validity == IntPtr.Zero)
            {
                return 0;
            }
            long unset = 0;
            for (long i = 0; i < length; i++)
            {
                byte b = Marshal.ReadByte(validity, (int)(i >> 3));
                if ((b & (1 << (int)(i & 7))) == 0)
                {
                    unset++;
                }
            }
            return unset;
        }

        private static IntPtr CopyBuffer(IntPtr source, long byteLength)
        {
            IntPtr destination = Marshal.AllocHGlobal((nint)Math.Max(byteLength, 1));
            if (byteLength > 0)
            {
                Buffer.MemoryCopy((void*)source, (void*)destination, byteLength, byteLength);
            }
            return destination;
        }

        private static IntPtr CopyOptionalBuffer(IntPtr source, long byteLength)
        {
            return source == IntPtr.Zero ? IntPtr.Zero : CopyBuffer(source, byteLength);
        }

        // xl_column's XL_T_BOOL is one byte per row (see NativeApi.Typed.cs); Arrow's canonical boolean
        // layout is bit-packed, LSB-first, same convention as a validity bitmap — this is the one
        // column type that cannot be a straight byte-for-byte copy.
        private static IntPtr BitPackBoolColumn(IntPtr byteValues, long length)
        {
            long byteLength = Math.Max((length + 7) / 8, 1);
            IntPtr destination = Marshal.AllocHGlobal((nint)byteLength);
            new Span<byte>((void*)destination, (int)byteLength).Clear();
            for (long i = 0; i < length; i++)
            {
                if (Marshal.ReadByte(byteValues, (int)i) != 0)
                {
                    byte* bytePtr = (byte*)destination + (i >> 3);
                    *bytePtr |= (byte)(1 << (int)(i & 7));
                }
            }
            return destination;
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
                _ => "tsu:", // NativeColumnType.Timestamp
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
            int columnSize = Marshal.SizeOf<NativeColumn>();
            return Marshal.PtrToStructure<NativeColumn>(IntPtr.Add(table.Columns, index * columnSize));
        }

        private static IntPtr SchemaReleaseCallback => (IntPtr)(delegate* unmanaged<ArrowSchema*, void>)&Exports.ReleaseArrowSchemaCallback;

        private static IntPtr ArrayReleaseCallback => (IntPtr)(delegate* unmanaged<ArrowArray*, void>)&Exports.ReleaseArrowArrayCallback;
    }
}
