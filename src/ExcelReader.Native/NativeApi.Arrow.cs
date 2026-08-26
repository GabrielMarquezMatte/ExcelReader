using System.Globalization;
using System.Numerics;
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
                // use. ReleaseArrowSchema takes an address, not a value, hence pinning it here.
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

            ReleaseChildren(schema.Children, schema.NChildren, &ReleaseArrowSchema);
            FreeIfSet(schema.Format);
            FreeIfSet(schema.Name);

            schema.Release = IntPtr.Zero;
            Marshal.StructureToPtr(schema, schemaPtr, false);
        }

        /// <summary>The <see cref="ArrowArray"/> half of <see cref="ReleaseArrowSchema"/>: the same
        /// recursive-children walk (shared via <see cref="ReleaseChildren"/>) and idempotent release
        /// mark, plus the buffers block, which a schema has no counterpart for.</summary>
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

        // Both release paths walk their children identically — release each, free each child's own
        // block, then free the child pointer array. Only the per-child release differs, so it arrives
        // as a function pointer, which in this unsafe context allocates no delegate.
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
                // Every child schema built before the throw (an AllocHGlobal OOM is the realistic
                // trigger - xl_parse_arrow briefly holds both the intermediate NativeTable and this
                // Arrow copy at once) is otherwise unreachable once this rethrows: it was never
                // returned to the caller, and ParseArrow's own catch only knows how to release a
                // *complete* schema via `schema`, which this method never got to assign. Release
                // exactly what got built here instead, the same way a real consumer eventually would.
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
                    NativeColumn column = ColumnAt(table, built);
                    IntPtr child = Marshal.AllocHGlobal(Marshal.SizeOf<ArrowArray>());
                    Marshal.StructureToPtr(BuildChildArray(specs[built].Type, column), child, false);
                    Marshal.WriteIntPtr(children, built * IntPtr.Size, child);
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
            catch
            {
                // See BuildArrowSchema's matching catch: every child array already copied out of
                // `table` before the throw is otherwise unreachable once this rethrows, since `array`
                // in ParseArrow is never assigned. Release exactly what got built.
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
            // PackBitsLsbFirst zero-fills the block and only ever sets bits below `length`, so every bit
            // past it is already 0 and popcount over whole bytes needs no tail mask.
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
            if (source == IntPtr.Zero)
            {
                return IntPtr.Zero;
            }
            return CopyBuffer(source, byteLength);
        }

        // xl_column's XL_T_BOOL is one byte per row (see NativeApi.Typed.cs); Arrow's canonical boolean
        // layout is bit-packed, LSB-first, the same convention as a validity bitmap — so this is the one
        // column type that cannot be a straight byte-for-byte copy, and the one place outside
        // ColumnBuilder that needs the shared packer.
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
