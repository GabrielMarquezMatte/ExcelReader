using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Tests.Reader.Internal
{
    public sealed class ReadOnlyMemoryStreamTests
    {
        [Fact]
        public async Task Reads_Seeks_And_Stops_At_The_End_Of_Memory_Without_An_Array()
        {
            using var manager = new NonArrayMemoryManager([1, 2, 3, 4, 5]);
            using var stream = new ReadOnlyMemoryStream(manager.Memory);
            byte[] buffer = new byte[3];

            Assert.Equal(5, stream.Length);
            Assert.Equal(3, stream.Read(buffer, 0, 3));
            Assert.Equal([1, 2, 3], buffer);
            Assert.Equal(4, stream.ReadByte());
            Assert.Equal(1, await stream.ReadAsync(buffer.AsMemory(), TestContext.Current.CancellationToken));
            Assert.Equal(5, buffer[0]);
            Assert.Equal(-1, stream.ReadByte());
            Assert.Equal(0, stream.Read(buffer));

            Assert.Equal(3, stream.Seek(-2, SeekOrigin.End));
            Assert.Equal(4, stream.ReadByte());
            Assert.Equal(1, stream.Seek(-3, SeekOrigin.Current));
            stream.Position = 10;
            Assert.Equal(0, stream.Read(buffer));
            Assert.Throws<IOException>(() => stream.Seek(-1, SeekOrigin.Begin));
            Assert.False(stream.CanWrite);
        }
    }
}
