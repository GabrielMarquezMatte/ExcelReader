using System.Buffers.Binary;
using System.Text;

namespace ExcelReader.Tests.Writer.Zip
{
    internal sealed record ZipCentral(
        int Position, ushort MadeBy, ushort Version, ushort Flags, ushort Method, uint Crc,
        uint CompressedSize, uint UncompressedSize, uint LocalOffset, string Name, byte[] Extra);

    internal sealed record ZipLocal(ushort Version, ushort Flags, ushort Method, uint Crc, uint CompressedSize, uint UncompressedSize);

    internal static class ZipLayout
    {
        internal static ushort U16(ReadOnlySpan<byte> data, long offset)
        {
            return BinaryPrimitives.ReadUInt16LittleEndian(data[(int)offset..]);
        }

        internal static uint U32(ReadOnlySpan<byte> data, long offset)
        {
            return BinaryPrimitives.ReadUInt32LittleEndian(data[(int)offset..]);
        }

        internal static ulong U64(ReadOnlySpan<byte> data, long offset)
        {
            return BinaryPrimitives.ReadUInt64LittleEndian(data[(int)offset..]);
        }

        /// <summary>Reads a classic (non-Zip64) central directory through the end record at the end of the file.</summary>
        internal static List<ZipCentral> ReadCentral(byte[] zip)
        {
            int end = zip.Length - 22;
            Assert.Equal(0x0605_4B50u, U32(zip, end));
            int count = U16(zip, end + 10);
            int position = (int)U32(zip, end + 16);
            var entries = new List<ZipCentral>(count);
            for (int i = 0; i < count; i++)
            {
                Assert.Equal(0x0201_4B50u, U32(zip, position));
                int nameLength = U16(zip, position + 28);
                int extraLength = U16(zip, position + 30);
                int commentLength = U16(zip, position + 32);
                entries.Add(new ZipCentral(
                    position, U16(zip, position + 4), U16(zip, position + 6), U16(zip, position + 8), U16(zip, position + 10),
                    U32(zip, position + 16), U32(zip, position + 20), U32(zip, position + 24), U32(zip, position + 42),
                    Encoding.UTF8.GetString(zip, position + 46, nameLength),
                    zip.AsSpan(position + 46 + nameLength, extraLength).ToArray()));
                position += 46 + nameLength + extraLength + commentLength;
            }
            return entries;
        }

        internal static ZipLocal ReadLocal(ReadOnlySpan<byte> zip, long offset)
        {
            Assert.Equal(0x0403_4B50u, U32(zip, offset));
            return new ZipLocal(U16(zip, offset + 4), U16(zip, offset + 6), U16(zip, offset + 8),
                U32(zip, offset + 14), U32(zip, offset + 18), U32(zip, offset + 22));
        }

        /// <summary>Zeroes the DOS time and date of every local and central header.</summary>
        internal static byte[] WithoutTimestamps(byte[] zip)
        {
            byte[] copy = (byte[])zip.Clone();
            foreach (ZipCentral entry in ReadCentral(zip))
            {
                copy.AsSpan((int)entry.LocalOffset + 10, 4).Clear();
                copy.AsSpan(entry.Position + 12, 4).Clear();
            }
            return copy;
        }
    }

    /// <summary>A write-only stream that refuses to seek, like a network or pipe stream.</summary>
    internal sealed class ForwardOnlyWriteStream(Stream inner) : Stream
    {
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush()
        {
            inner.Flush();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotSupportedException();
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            inner.Write(buffer, offset, count);
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            inner.Write(buffer);
        }
    }

    /// <summary>Accepts writes until <paramref name="limit"/> bytes, then throws <see cref="IOException"/> on every write.</summary>
    internal sealed class FailingWriteStream(long limit) : Stream
    {
        internal MemoryStream Written { get; } = new();
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotSupportedException();
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            Write(buffer.AsSpan(offset, count));
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            if (Written.Length + buffer.Length > limit)
            {
                throw new IOException("disk full");
            }
            Written.Write(buffer);
        }
    }

    /// <summary>Throws <see cref="IOException"/> on write number <paramref name="failOnCall"/> without storing it, then accepts every later write.</summary>
    internal sealed class FailOnceWriteStream(int failOnCall) : Stream
    {
        private int _calls;

        internal MemoryStream Written { get; } = new();
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotSupportedException();
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            Write(buffer.AsSpan(offset, count));
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            _calls++;
            if (_calls == failOnCall)
            {
                throw new IOException("transient");
            }
            Written.Write(buffer);
        }
    }

    /// <summary>
    /// Counts every byte but keeps only the first 4 KB and everything at or past 4 GiB, so a Zip64
    /// archive can be inspected without holding it.
    /// </summary>
    internal sealed class CapturingStream(bool seekable) : Stream
    {
        internal const long TailStart = 1L << 32;
        private readonly MemoryStream _tail = new();
        private long _position;
        private long _length;

        internal byte[] Head { get; } = new byte[4096];
        internal byte[] Tail => _tail.ToArray();
        public override bool CanRead => false;
        public override bool CanSeek => seekable;
        public override bool CanWrite => true;
        public override long Length => _length;

        public override long Position
        {
            get => _position;
            set => Seek(value, SeekOrigin.Begin);
        }

        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            Write(buffer.AsSpan(offset, count));
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            if (!seekable || origin != SeekOrigin.Begin)
            {
                throw new NotSupportedException();
            }
            _position = offset;
            return _position;
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            long start = _position;
            if (start < Head.Length)
            {
                int headPart = (int)Math.Min(buffer.Length, Head.Length - start);
                buffer[..headPart].CopyTo(Head.AsSpan((int)start));
            }
            long end = start + buffer.Length;
            if (end > TailStart)
            {
                long from = Math.Max(start, TailStart);
                _tail.Position = from - TailStart;
                _tail.Write(buffer[(int)(from - start)..]);
            }
            _position = end;
            _length = Math.Max(_length, end);
        }
    }
}
