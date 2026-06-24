using System.Buffers;
using System.Buffers.Binary;

namespace ExcelReader.Core.Writer.Internal
{
    // Growable, pooled little-endian byte buffer for assembling a BIFF substream.
    // Record framing: BeginRecord writes the id + a placeholder length; EndRecord back-patches
    // the length once the payload is written, so records can be built field-by-field in place.
    internal sealed class BiffBuffer : IDisposable
    {
        private byte[] _buffer;
        private int _length;

        internal BiffBuffer(int initialCapacity = 4096)
        {
            _buffer = ArrayPool<byte>.Shared.Rent(initialCapacity);
        }

        internal int Length => _length;

        internal ReadOnlySpan<byte> Span => _buffer.AsSpan(0, _length);

        internal ReadOnlyMemory<byte> Memory => _buffer.AsMemory(0, _length);

        internal void WriteByte(byte value)
        {
            Ensure(1);
            _buffer[_length++] = value;
        }

        internal void WriteU16(int value)
        {
            Ensure(2);
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(_length), (ushort)value);
            _length += 2;
        }

        internal void WriteI32(int value)
        {
            Ensure(4);
            BinaryPrimitives.WriteInt32LittleEndian(_buffer.AsSpan(_length), value);
            _length += 4;
        }

        internal void WriteU32(uint value)
        {
            Ensure(4);
            BinaryPrimitives.WriteUInt32LittleEndian(_buffer.AsSpan(_length), value);
            _length += 4;
        }

        internal void WriteDouble(double value)
        {
            Ensure(8);
            BinaryPrimitives.WriteDoubleLittleEndian(_buffer.AsSpan(_length), value);
            _length += 8;
        }

        internal void Write(ReadOnlySpan<byte> bytes)
        {
            Ensure(bytes.Length);
            bytes.CopyTo(_buffer.AsSpan(_length));
            _length += bytes.Length;
        }

        // Writes the record id + a 2-byte length placeholder; returns the placeholder offset.
        internal int BeginRecord(int id)
        {
            WriteU16(id);
            int lengthPos = _length;
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
            int payload = _length - lengthPos - 2;
            if ((uint)payload > BiffRecord.MaxPayload)
            {
                throw new InvalidOperationException($"BIFF record payload {payload} exceeds the {BiffRecord.MaxPayload}-byte limit.");
            }
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.AsSpan(lengthPos), (ushort)payload);
        }

        private void Ensure(int extra)
        {
            int needed = _length + extra;
            if (needed <= _buffer.Length)
            {
                return;
            }
            byte[] bigger = ArrayPool<byte>.Shared.Rent(Math.Max(_buffer.Length * 2, needed));
            _buffer.AsSpan(0, _length).CopyTo(bigger);
            ArrayPool<byte>.Shared.Return(_buffer);
            _buffer = bigger;
        }

        public void Dispose()
        {
            if (_buffer.Length > 0)
            {
                ArrayPool<byte>.Shared.Return(_buffer);
                _buffer = [];
                _length = 0;
            }
        }
    }
}
