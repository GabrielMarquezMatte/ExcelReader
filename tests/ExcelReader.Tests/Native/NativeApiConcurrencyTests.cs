using System.Text;
using ExcelReader.Native;
using ExcelReader.Native.Reading;

namespace ExcelReader.Tests.Native
{
    public class NativeApiConcurrencyTests
    {
        private static readonly string XlsxFixture = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");

        [Fact]
        public Task IndependentHandlesOnSeparateThreadsDoNotInterfere()
        {
            const int handleCount = 16;
            IEnumerable<Task> tasks = Enumerable.Range(0, handleCount).Select(_ => Task.Run(() =>
            {
                Assert.Equal(NativeStatus.Ok, ReadApi.OpenFile(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Auto, out NativeHandle? handle));
                Assert.NotNull(handle);

                int rowCount = 0;
                while (ReadApi.NextRow(handle, new byte[64 * 1024], out _) != NativeStatus.Eof)
                {
                    rowCount++;
                }

                Assert.Equal(3, rowCount);
                Assert.Equal(NativeStatus.Ok, ReadApi.Close(handle));
            }));

            return Task.WhenAll(tasks);
        }

        [Fact]
        public Task LastErrorIsPerThreadNotSharedAcrossHandles()
        {
            const int threadCount = 16;
            IEnumerable<Task> tasks = Enumerable.Range(0, threadCount).Select(i => Task.Run(() =>
            {
                if (i % 2 == 0)
                {
                    byte[] missingPath = Encoding.UTF8.GetBytes($"does-not-exist-{i}.xlsx");
                    int status = ReadApi.OpenFile(missingPath, NativeFormat.Auto, out NativeHandle? handle);
                    Assert.NotEqual(NativeStatus.Ok, status);
                    Assert.Null(handle);

                    Span<byte> errorBuffer = stackalloc byte[256];
                    Assert.Equal(NativeStatus.Ok, NativeApi.LastError(errorBuffer, out int length));
                    Assert.True(length > 0);
                }
                else
                {
                    Assert.Equal(NativeStatus.Ok, ReadApi.OpenFile(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Auto, out NativeHandle? handle));
                    Assert.NotNull(handle);

                    Span<byte> errorBuffer = stackalloc byte[256];
                    Assert.Equal(NativeStatus.Ok, NativeApi.LastError(errorBuffer, out int length));
                    Assert.Equal(0, length);

                    Assert.Equal(NativeStatus.Ok, ReadApi.Close(handle));
                }
            }));

            return Task.WhenAll(tasks);
        }
    }
}
