using System.Buffers;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader.Schema;

namespace ExcelReader.Core.Reader.Columnar
{
    internal sealed unsafe partial class TypedColumnBuilder
    {
        /// <summary>Joins one column's parts, in order, into a single column.</summary>
        internal static BuiltColumn Build(ReadOnlySpan<TypedColumnBuilder> parts)
        {
            ExcelColumnType type = parts[0].Type;
            long rows = 0;
            bool anyNull = false;
            foreach (TypedColumnBuilder part in parts)
            {
                rows += part._rowCount;
                anyNull |= part._anyNull;
            }
            BuiltColumn column = new() { Type = type, Length = rows };
            try
            {
                if (anyNull)
                {
                    column.Validity = ConcatValidity(parts, rows);
                }
                FillValues(ref column, parts);
            }
            catch
            {
                column.Free();
                throw;
            }
            return column;
        }

        private static void FillValues(ref BuiltColumn column, ReadOnlySpan<TypedColumnBuilder> parts)
        {
            switch (column.Type)
            {
                case ExcelColumnType.StringColumn:
                    column.Values = ConcatStringOffsets(parts);
                    column.Data = Concat(parts, static part => part._stringData, out int dataLength);
                    column.DataLen = dataLength;
                    return;
                case ExcelColumnType.StringDictionaryColumn:
                    FillDictionary(ref column, parts);
                    return;
                case ExcelColumnType.BoolColumn:
                    column.Values = Concat(parts, static part => part._bools, out _);
                    return;
                case ExcelColumnType.Float64Column:
                    column.Values = Concat(parts, static part => part._doubles, out _);
                    return;
                case ExcelColumnType.DateColumn:
                    column.Values = Concat(parts, static part => part._ints, out _);
                    return;
                default:
                    column.Values = Concat(parts, static part => part._longs, out _);
                    return;
            }
        }

        private static void FillDictionary(ref BuiltColumn column, ReadOnlySpan<TypedColumnBuilder> parts)
        {
            StringDictionary merged = parts[0]._dictionary!;
            column.Values = parts.Length == 1 ? Concat(parts, static part => part._ints, out _) : ConcatRemappedIndices(parts, merged);
            column.DictionaryLength = merged.Count;
            column.DictionaryOffsets = Finish(merged.Offsets, out _);
            column.Data = Finish(merged.Data, out int dataLength);
            column.DataLen = dataLength;
        }

        private static IntPtr ConcatRemappedIndices(ReadOnlySpan<TypedColumnBuilder> parts, StringDictionary merged)
        {
            long rows = 0;
            foreach (TypedColumnBuilder part in parts)
            {
                rows += part._ints.Count;
            }
            IntPtr block = ColumnBlocks.Alloc(Math.Max(checked((int)rows * sizeof(int)), 1));
            Span<int> indices = new((void*)block, (int)rows);
            parts[0]._ints.AsSpan().CopyTo(indices);
            int at = parts[0]._ints.Count;
            for (int p = 1; p < parts.Length; p++)
            {
                StringDictionary dictionary = parts[p]._dictionary!;
                int[] remap = new int[Math.Max(dictionary.Count, 1)];
                for (int entry = 0; entry < dictionary.Count; entry++)
                {
                    remap[entry] = merged.GetOrAdd(dictionary.ValueAt(entry));
                }
                foreach (int index in parts[p]._ints.AsSpan())
                {
                    indices[at++] = remap[index];
                }
            }
            return block;
        }

        private static IntPtr Concat<T>(ReadOnlySpan<TypedColumnBuilder> parts, Func<TypedColumnBuilder, GrowableNativeBuffer<T>> select, out int byteLength)
            where T : unmanaged
        {
            if (parts.Length == 1)
            {
                return Finish(select(parts[0]), out byteLength);
            }
            byteLength = 0;
            foreach (TypedColumnBuilder part in parts)
            {
                byteLength = checked(byteLength + select(part).ByteLength);
            }
            IntPtr block = ColumnBlocks.Alloc(Math.Max(byteLength, 1));
            int offset = 0;
            foreach (TypedColumnBuilder part in parts)
            {
                GrowableNativeBuffer<T> buffer = select(part);
                buffer.CopyTo(new Span<byte>((byte*)block + offset, buffer.ByteLength));
                offset += buffer.ByteLength;
            }
            return block;
        }

        // A buffer's block already is the column, unless a test hook is tracking allocations.
        private static IntPtr Finish<T>(GrowableNativeBuffer<T> buffer, out int byteLength) where T : unmanaged
        {
            byteLength = buffer.ByteLength;
            if (ColumnBlocks.AllocOverride is null && buffer.Count > 0)
            {
                return buffer.Detach();
            }
            IntPtr block = ColumnBlocks.Alloc(Math.Max(byteLength, 1));
            buffer.CopyTo(new Span<byte>((void*)block, byteLength));
            return block;
        }

        private static IntPtr ConcatStringOffsets(ReadOnlySpan<TypedColumnBuilder> parts)
        {
            if (parts.Length == 1)
            {
                return Finish(parts[0]._stringOffsets, out _);
            }
            int count = 1;
            foreach (TypedColumnBuilder part in parts)
            {
                count = checked(count + part._stringOffsets.Count - 1);
            }
            IntPtr block = ColumnBlocks.Alloc(checked(count * sizeof(int)));
            Span<int> offsets = new((void*)block, count);
            offsets[0] = 0;
            int at = 0;
            foreach (TypedColumnBuilder part in parts)
            {
                int baseOffset = offsets[at];
                Span<int> slot = offsets.Slice(at, part._stringOffsets.Count);
                part._stringOffsets.CopyTo(MemoryMarshal.AsBytes(slot));
                if (baseOffset != 0)
                {
                    foreach (ref int offset in slot)
                    {
                        offset += baseOffset;
                    }
                }
                at += slot.Length - 1;
            }
            return block;
        }

        private static IntPtr ConcatValidity(ReadOnlySpan<TypedColumnBuilder> parts, long rows)
        {
            int byteLength = checked((int)((rows + 7) / 8));
            IntPtr block = ColumnBlocks.Alloc(Math.Max(byteLength, 1));
            Span<byte> bits = new((void*)block, byteLength);
            bits.Clear();
            long at = 0;
            foreach (TypedColumnBuilder part in parts)
            {
                WritePartBits(bits, at, part);
                at += part._rowCount;
            }
            return block;
        }

        private static void WritePartBits(Span<byte> bits, long at, TypedColumnBuilder part)
        {
            if (!part._anyNull)
            {
                SetBits(bits, at, part._rowCount);
                return;
            }
            OrBits(bits, at, part._validity);
        }

        private static void OrBits(Span<byte> bits, long at, GrowableNativeBuffer<byte> validity)
        {
            int first = (int)(at >> 3);
            int shift = (int)(at & 7);
            if (shift == 0)
            {
                validity.CopyTo(bits[first..]);
                return;
            }
            int count = validity.Count;
            byte[] source = ArrayPool<byte>.Shared.Rent(count);
            try
            {
                validity.CopyTo(source);
                for (int i = 0; i < count; i++)
                {
                    bits[first + i] |= (byte)(source[i] << shift);
                    if (first + i + 1 < bits.Length)
                    {
                        bits[first + i + 1] |= (byte)(source[i] >> (8 - shift));
                    }
                }
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(source);
            }
        }

        private static void SetBits(Span<byte> bits, long start, int count)
        {
            long end = start + count;
            for (; start < end && (start & 7) != 0; start++)
            {
                bits[(int)(start >> 3)] |= (byte)(1 << (int)(start & 7));
            }
            int fullBytes = (int)((end - start) >> 3);
            bits.Slice((int)(start >> 3), fullBytes).Fill(0xFF);
            for (start += (long)fullBytes << 3; start < end; start++)
            {
                bits[(int)(start >> 3)] |= (byte)(1 << (int)(start & 7));
            }
        }
    }
}
