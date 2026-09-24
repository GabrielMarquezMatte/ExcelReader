using System.Buffers;
using System.Buffers.Binary;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Writer.Xls;

namespace ExcelReader.Core.Writer.Internal
{
    internal sealed class BiffBuffer : IDisposable
    {
        // ponytail: 32 MB cap; sheets past that fall back to plain allocs (same as Shared did).
        private static readonly ArrayPool<byte> DefaultPool = ArrayPool<byte>.Create(32 * 1024 * 1024, 16);
        private static readonly AsyncLocal<ArrayPool<byte>?> PoolOverride = new();

        private static ArrayPool<byte> Pool => PoolOverride.Value ?? DefaultPool;

        // Test seam: lets a test count every rent/return made on its own async flow.
        internal static void OverridePool(ArrayPool<byte>? pool)
        {
            PoolOverride.Value = pool;
        }

        private byte[] _buffer;

        internal BiffBuffer(int initialCapacity = 4096)
        {
            _buffer = Pool.Rent(initialCapacity);
        }

        internal int Length { get; private set; }

        internal bool IsReleased => _buffer.Length == 0;

        internal ReadOnlySpan<byte> Span => _buffer.AsSpan(0, Length);

        internal ReadOnlyMemory<byte> Memory => _buffer.AsMemory(0, Length);

        internal void Reset()
        {
            Length = 0;
        }

        internal byte[] Detach(out int length)
        {
            byte[] detached = _buffer;
            length = Length;
            _buffer = Pool.Rent(detached.Length);
            Length = 0;
            return detached;
        }

        internal static void ReturnDetached(byte[] buffer)
        {
            if (buffer.Length > 0)
            {
                Pool.Return(buffer);
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void WriteByte(byte value)
        {
            Ensure(1);
            _buffer[Length++] = value;
        }

        internal void WriteU16(int value)
        {
            Ensure(2);
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(Length), (ushort)value);
            Length += 2;
        }

        internal void WriteI32(int value)
        {
            Ensure(4);
            BinaryPrimitives.WriteInt32LittleEndian(_buffer.AsSpan(Length), value);
            Length += 4;
        }

        internal void WriteU32(uint value)
        {
            Ensure(4);
            BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(Length), value);
            Length += 4;
        }

        internal void WriteDouble(double value)
        {
            Ensure(8);
            BinaryPrimitives.WriteDoubleLittleEndian(_buffer.AsSpan(Length), value);
            Length += 8;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void Write(ReadOnlySpan<byte> bytes)
        {
            Ensure(bytes.Length);
            bytes.CopyTo(_buffer.AsSpan(Length));
            Length += bytes.Length;
        }

        internal Span<byte> GetSpan(int sizeHint)
        {
            Ensure(sizeHint);
            return _buffer.AsSpan(Length);
        }

        internal void Advance(int count)
        {
            Length += count;
        }

        internal void WriteUtf8(ReadOnlySpan<char> chars)
        {
            Ensure(checked(chars.Length * 3));
            Length += Encoding.UTF8.GetBytes(chars, _buffer.AsSpan(Length));
        }

        internal void WriteUtf16(ReadOnlySpan<char> chars)
        {
            int byteCount = checked(chars.Length * sizeof(char));
            Ensure(byteCount);
            MemoryMarshal.AsBytes(chars).CopyTo(_buffer.AsSpan(Length, byteCount));
            Length += byteCount;
        }

        internal Span<byte> Slice(int position, int length)
        {
            return _buffer.AsSpan(position, length);
        }

        internal void Truncate(int length)
        {
            Length = length;
        }

        internal int BeginRecord(int id)
        {
            WriteU16(id);
            int lengthPos = Length;
            WriteU16(0);
            return lengthPos;
        }

        internal void PatchI32(int position, int value)
        {
            BinaryPrimitives.WriteInt32LittleEndian(_buffer.AsSpan(position, 4), value);
        }

        internal void EndRecord(int lengthPos)
        {
            int payload = Length - lengthPos - 2;
            if ((uint)payload > BiffRecord.MaxPayload)
            {
                throw new InvalidOperationException($"BIFF record payload {payload} exceeds the {BiffRecord.MaxPayload}-byte limit.");
            }
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(lengthPos), (ushort)payload);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void Ensure(int extra)
        {
            int needed = Length + extra;
            if (needed > _buffer.Length)
            {
                Grow(needed);
            }
        }

        // Kept out of line so the inlined Ensure on every Write stays a single compare.
        [MethodImpl(MethodImplOptions.NoInlining)]
        private void Grow(int needed)
        {
            byte[] bigger = Pool.Rent(Math.Max(_buffer.Length * 2, needed));
            _buffer.AsSpan(0, Length).CopyTo(bigger);
            Pool.Return(_buffer);
            _buffer = bigger;
        }

        public void Dispose()
        {
            if (_buffer.Length > 0)
            {
                Pool.Return(_buffer);
                _buffer = [];
                Length = 0;
            }
        }
    }
}
