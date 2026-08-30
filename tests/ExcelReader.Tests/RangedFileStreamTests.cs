using System.Runtime.InteropServices;
using ExcelReader.Core.Parser.Internal;
using Microsoft.Testing.Platform.Services;
using Microsoft.Win32.SafeHandles;

namespace ExcelReader.Tests
{
    public class RangedFileStreamTests
    {
        private static string WriteTemp(byte[] content)
        {
            string path = Path.Combine(Path.GetTempPath(), $"exr-{Guid.NewGuid():N}.bin");
            File.WriteAllBytes(path, content);
            return path;
        }

        [Fact]
        public void ReadsFromTheStartOffsetToEndOfFile()
        {
            string path = WriteTemp("0123456789"u8.ToArray());
            try
            {
                using SafeFileHandle handle = File.OpenHandle(path);
                using var stream = new RangedFileStream(handle, start: 4);

                byte[] buffer = new byte[16];
                int n = stream.Read(buffer, 0, buffer.Length);

                Assert.Equal(6, n);
                Assert.Equal("456789"u8.ToArray(), buffer.AsSpan(0, n).ToArray());
                Assert.Equal(0, stream.Read(buffer, 0, buffer.Length));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public async Task ReadAsyncAdvancesTheOffsetAcrossCalls()
        {
            string path = WriteTemp("abcdefgh"u8.ToArray());
            try
            {
                using SafeFileHandle handle = File.OpenHandle(path, options: FileOptions.Asynchronous);
                await using var stream = new RangedFileStream(handle, start: 2);

                byte[] first = new byte[3];
                byte[] second = new byte[3];
                int n1 = await stream.ReadAsync(first, TestContext.Current.CancellationToken);
                int n2 = await stream.ReadAsync(second, TestContext.Current.CancellationToken);

                Assert.Equal(3, n1);
                Assert.Equal("cde"u8.ToArray(), first);
                Assert.Equal(3, n2);
                Assert.Equal("fgh"u8.ToArray(), second);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void TwoStreamsShareOneHandleWithoutInterferingWithEachOther()
        {
            string path = WriteTemp("ABCDEFGHIJ"u8.ToArray());
            try
            {
                using SafeFileHandle handle = File.OpenHandle(path);
                using var a = new RangedFileStream(handle, start: 0);
                using var b = new RangedFileStream(handle, start: 5);

                byte[] bufA = new byte[3];
                byte[] bufB = new byte[3];
                a.Read(bufA, 0, 3);
                b.Read(bufB, 0, 3);
                byte[] bufA2 = new byte[2];
                a.Read(bufA2, 0, 2);

                Assert.Equal("ABC"u8.ToArray(), bufA);
                Assert.Equal("FGH"u8.ToArray(), bufB);
                Assert.Equal("DE"u8.ToArray(), bufA2);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void DisposingTheStreamLeavesTheSharedHandleOpen()
        {
            string path = WriteTemp("xyz"u8.ToArray());
            try
            {
                using SafeFileHandle handle = File.OpenHandle(path);
                using (var stream = new RangedFileStream(handle, start: 0))
                {
                    _ = stream.ReadByte();
                }

                Assert.False(handle.IsClosed);
                using var again = new RangedFileStream(handle, start: 0);
                Assert.Equal((byte)'x', again.ReadByte());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void SeekingIsNotSupported()
        {
            string path = WriteTemp("q"u8.ToArray());
            try
            {
                using SafeFileHandle handle = File.OpenHandle(path);
                using var stream = new RangedFileStream(handle, start: 0);

                Assert.False(stream.CanSeek);
                Assert.Throws<NotSupportedException>(() => stream.Seek(0, SeekOrigin.Begin));
                Assert.Throws<NotSupportedException>(() => stream.Length);
                Assert.Throws<NotSupportedException>(() => stream.Position);
            }
            finally
            {
                File.Delete(path);
            }
        }
    }
}
