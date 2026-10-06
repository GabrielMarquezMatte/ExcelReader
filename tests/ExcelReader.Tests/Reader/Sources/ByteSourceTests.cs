using ExcelReader.Core.Reader.Sources;

namespace ExcelReader.Tests.Reader.Sources
{
    public sealed class ByteSourceTests : IDisposable
    {
        public enum Kind
        {
            Memory,
            ExposedMemoryStream,
            OpaqueMemoryStream,
            Trickle,
            FileStream,
            FilePath,
            NonSeekable,
        }

        private static readonly byte[] Payload = BuildPayload(200_000);

        private readonly string _path = Path.Combine(Path.GetTempPath(), "excelreader-bytesource-" + Guid.NewGuid().ToString("N") + ".bin");

        public ByteSourceTests()
        {
            File.WriteAllBytes(_path, Payload);
        }

        public void Dispose()
        {
            File.Delete(_path);
        }

        public static TheoryData<Kind> Kinds =>
        [
            Kind.Memory, Kind.ExposedMemoryStream, Kind.OpaqueMemoryStream, Kind.Trickle,
            Kind.FileStream, Kind.FilePath, Kind.NonSeekable,
        ];

        private static byte[] BuildPayload(int length)
        {
            byte[] bytes = new byte[length];
            uint state = 12345;
            for (int i = 0; i < length; i++)
            {
                state = (state * 1664525) + 1013904223;
                bytes[i] = (byte)(state >> 24);
            }
            return bytes;
        }

        private ByteSource Open(Kind kind)
        {
            switch (kind)
            {
                case Kind.Memory:
                    return ByteSource.FromMemory(Payload);
                case Kind.ExposedMemoryStream:
                    MemoryStream exposed = new();
                    exposed.Write(Payload);
                    exposed.Position = 17;
                    return ByteSource.FromStream(exposed, leaveOpen: false);
                case Kind.OpaqueMemoryStream:
                    return ByteSource.FromStream(new MemoryStream(Payload, writable: false), leaveOpen: false);
                case Kind.Trickle:
                    return ByteSource.FromStream(new TrickleStream(Payload), leaveOpen: false);
                case Kind.FileStream:
                    return ByteSource.FromStream(File.OpenRead(_path), leaveOpen: false);
                case Kind.FilePath:
                    return ByteSource.OpenFile(_path, asynchronous: false);
                default:
                    return ByteSource.FromStream(new NonSeekableStream(Payload), leaveOpen: false);
            }
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void Length_Is_The_Whole_Payload(Kind kind)
        {
            using ByteSource source = Open(kind);
            Assert.Equal(Payload.Length, source.Length);
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void ReadExactly_Returns_The_Bytes_At_The_Offset(Kind kind)
        {
            using ByteSource source = Open(kind);
            byte[] read = new byte[5000];
            source.ReadExactly(150_000, read);
            Assert.True(read.AsSpan().SequenceEqual(Payload.AsSpan(150_000, 5000)));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public async Task ReadExactlyAsync_Returns_The_Bytes_At_The_Offset(Kind kind)
        {
            using ByteSource source = Open(kind);
            byte[] read = new byte[5000];
            await source.ReadExactlyAsync(150_000, read, TestContext.Current.CancellationToken);
            Assert.True(read.AsSpan().SequenceEqual(Payload.AsSpan(150_000, 5000)));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void Read_At_The_End_Returns_Zero(Kind kind)
        {
            using ByteSource source = Open(kind);
            Assert.Equal(0, source.Read(Payload.Length, new byte[16]));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void ReadExactly_Past_The_End_Throws(Kind kind)
        {
            using ByteSource source = Open(kind);
            Assert.Throws<EndOfStreamException>(() => source.ReadExactly(Payload.Length - 10, new byte[11]));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void TryGetMemory_Is_True_Only_When_The_Bytes_Are_In_Memory(Kind kind)
        {
            using ByteSource source = Open(kind);
            bool expected = kind is Kind.Memory or Kind.ExposedMemoryStream or Kind.NonSeekable;
            Assert.Equal(expected, source.TryGetMemory(out ReadOnlyMemory<byte> memory));
            if (expected)
            {
                Assert.True(memory.Span.SequenceEqual(Payload));
            }
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void Concurrent_Reads_Each_Get_Their_Own_Bytes(Kind kind)
        {
            using ByteSource source = Open(kind);
            int iterations = kind == Kind.Trickle ? 20 : 200;
            Parallel.For(0, 16, worker =>
            {
                uint state = (uint)worker + 1;
                byte[] buffer = new byte[4096];
                for (int i = 0; i < iterations; i++)
                {
                    state = (state * 1664525) + 1013904223;
                    int offset = (int)(state % (uint)(Payload.Length - buffer.Length));
                    source.ReadExactly(offset, buffer);
                    Assert.True(buffer.AsSpan().SequenceEqual(Payload.AsSpan(offset, buffer.Length)));
                }
            });
        }

        [Fact]
        public void Dispose_Closes_A_Stream_It_Owns()
        {
            TrickleStream stream = new(Payload);
            ByteSource.FromStream(stream, leaveOpen: false).Dispose();
            Assert.False(stream.CanRead);
        }

        [Fact]
        public void Dispose_Leaves_A_Borrowed_Stream_Open()
        {
            using TrickleStream stream = new(Payload);
            ByteSource.FromStream(stream, leaveOpen: true).Dispose();
            Assert.True(stream.CanRead);
        }

        [Fact]
        public void A_Borrowed_FileStream_Stays_Usable()
        {
            using FileStream stream = File.OpenRead(_path);
            ByteSource.FromStream(stream, leaveOpen: true).Dispose();
            Assert.Equal(Payload[0], stream.ReadByte());
        }

        [Fact]
        public void Dispose_Releases_The_File_Handle()
        {
            ByteSource.OpenFile(_path, asynchronous: false).Dispose();
            File.Delete(_path);
            Assert.False(File.Exists(_path));
        }

        [Fact]
        public async Task FromStreamAsync_Buffers_A_NonSeekable_Stream()
        {
            using ByteSource source = await ByteSource.FromStreamAsync(
                new NonSeekableStream(Payload), leaveOpen: false, TestContext.Current.CancellationToken);
            Assert.True(source.TryGetMemory(out ReadOnlyMemory<byte> memory));
            Assert.True(memory.Span.SequenceEqual(Payload));
        }

        private sealed class InvertingMemoryStream(byte[] bytes) : MemoryStream(bytes, 0, bytes.Length, writable: false, publiclyVisible: true)
        {
            public override int Read(byte[] buffer, int offset, int count)
            {
                int read = base.Read(buffer, offset, count);
                for (int i = 0; i < read; i++)
                {
                    buffer[offset + i] ^= 0xFF;
                }
                return read;
            }

            public override int Read(Span<byte> buffer)
            {
                byte[] scratch = new byte[buffer.Length];
                int read = Read(scratch, 0, scratch.Length);
                scratch.AsSpan(0, read).CopyTo(buffer);
                return read;
            }
        }

        [Fact]
        public void A_MemoryStream_Subclass_Is_Read_Through_Its_Own_Read()
        {
            using ByteSource source = ByteSource.FromStream(new InvertingMemoryStream(Payload), leaveOpen: false);
            Assert.False(source.TryGetMemory(out _));
            byte[] read = new byte[64];
            source.ReadExactly(100, read);
            for (int i = 0; i < read.Length; i++)
            {
                Assert.Equal((byte)(Payload[100 + i] ^ 0xFF), read[i]);
            }
        }
    }
}
