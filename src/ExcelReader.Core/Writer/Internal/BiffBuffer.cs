using System.Buffers;
using System.Buffers.Binary;
using System.Text;

namespace ExcelReader.Core.Writer.Internal
{
    // Growable, pooled little-endian byte buffer for assembling a BIFF substream.
    // Record framing: BeginRecord writes the id + a placeholder length; EndRecord back-patches
    // the length once the payload is written, so records can be built field-by-field in place.
    internal sealed class BiffBuffer : IDisposable
    {
        // A single sheet's cell buffer routinely exceeds ArrayPool.Shared's 1 MB cap (a 50k-row
        // sheet is ~4 MB), so those rents would never be recycled — every doubling leaks to the GC
        // as LOH garbage. A dedicated pool with a higher cap reuses them across sheets/workbooks.
        // ponytail: 32 MB cap; sheets past that fall back to plain allocs (same as Shared did).
        private static readonly ArrayPool<byte> Pool = ArrayPool<byte>.Create(32 * 1024 * 1024, 16);

        private byte[] _buffer;

        internal BiffBuffer(int initialCapacity = 4096)
        {
            _buffer = Pool.Rent(initialCapacity);
        }

        internal int Length { get; private set; }

        internal ReadOnlySpan<byte> Span => _buffer.AsSpan(0, Length);

        internal ReadOnlyMemory<byte> Memory => _buffer.AsMemory(0, Length);

        // Rewinds to empty, keeping the rented buffer for reuse.
        internal void Reset()
        {
            Length = 0;
        }

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

        internal void Write(ReadOnlySpan<byte> bytes)
        {
            Ensure(bytes.Length);
            bytes.CopyTo(_buffer.AsSpan(Length));
            Length += bytes.Length;
        }

        // Reserves at least `sizeHint` free bytes and hands back the writable tail so callers can
        // format directly into the buffer (e.g. Utf8Formatter), then commit with Advance — saves the
        // temp-span + copy that Write would otherwise need.
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
            // Single pass: reserve the UTF-8 worst case (3 bytes per UTF-16 code unit) and encode once.
            Ensure(checked(chars.Length * 3));
            Length += Encoding.UTF8.GetBytes(chars, _buffer.AsSpan(Length));
        }

        internal void WriteUtf16(ReadOnlySpan<char> chars)
        {
            int byteCount = checked(chars.Length * sizeof(char));
            Ensure(byteCount);
            Span<byte> dest = _buffer.AsSpan(Length, byteCount);
            for (int i = 0; i < chars.Length; i++)
            {
                BinaryPrimitives.WriteUInt16LittleEndian(dest[(i * sizeof(char))..], chars[i]);
            }
            Length += byteCount;
        }

        // Writes the record id + a 2-byte length placeholder; returns the placeholder offset.
        internal int BeginRecord(int id)
        {
            WriteU16(id);
            int lengthPos = Length;
            WriteU16(0);
            return lengthPos;
        }

        // Overwrites a previously written 32-bit value (e.g. a BoundSheet offset placeholder).
        internal void PatchI32(int position, int value)
        {
            BinaryPrimitives.WriteInt32LittleEndian(_buffer.AsSpan(position, 4), value);
        }

        // Back-patches the length written by BeginRecord with the actual payload size.
        internal void EndRecord(int lengthPos)
        {
            int payload = Length - lengthPos - 2;
            if ((uint)payload > BiffRecord.MaxPayload)
            {
                throw new InvalidOperationException($"BIFF record payload {payload} exceeds the {BiffRecord.MaxPayload}-byte limit.");
            }
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(lengthPos), (ushort)payload);
        }

        private void Ensure(int extra)
        {
            int needed = Length + extra;
            if (needed <= _buffer.Length)
            {
                return;
            }
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
