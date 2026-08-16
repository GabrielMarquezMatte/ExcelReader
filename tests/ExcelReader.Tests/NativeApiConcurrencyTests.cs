using System.Text;
using ExcelReader.Native;

namespace ExcelReader.Tests
{
    // Covers the concurrency contract the header documents: one handle per thread. Two independent
    // handles used one-per-thread must not interfere with each other or see each other's errors —
    // NOT that a single handle may be shared across threads, which is explicitly unsupported and
    // stays that way. Drives NativeApi directly: the [UnmanagedCallersOnly] Exports entry points
    // cannot be called from managed code.
    public class NativeApiConcurrencyTests
    {
        private static readonly string XlsxFixture = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");

        [Fact]
        public Task IndependentHandlesOnSeparateThreadsDoNotInterfere()
        {
            const int handleCount = 16;
            IEnumerable<Task> tasks = Enumerable.Range(0, handleCount).Select(_ => Task.Run(() =>
            {
                Assert.Equal(NativeStatus.Ok, NativeApi.OpenFile(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Auto, out NativeHandle? handle));
                Assert.NotNull(handle);

                int rowCount = 0;
                while (NativeApi.NextRow(handle, new byte[64 * 1024], out _) != NativeStatus.Eof)
                {
                    rowCount++;
                }

                Assert.Equal(3, rowCount); // 1 header + 2 data rows — see SampleTest for this fixture's actual shape
                Assert.Equal(NativeStatus.Ok, NativeApi.Close(handle));
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
                    // Odd/even split: half the threads provoke a real error (missing file), half do a
                    // normal open. A LastError leaking across threads would show up as the "clean" half
                    // seeing the other half's error message, or vice versa.
                    byte[] missingPath = Encoding.UTF8.GetBytes($"does-not-exist-{i}.xlsx");
                    int status = NativeApi.OpenFile(missingPath, NativeFormat.Auto, out NativeHandle? handle);
                    Assert.NotEqual(NativeStatus.Ok, status);
                    Assert.Null(handle);

                    Span<byte> errorBuffer = stackalloc byte[256];
                    Assert.Equal(NativeStatus.Ok, NativeApi.LastError(errorBuffer, out int length));
                    Assert.True(length > 0);
                }
                else
                {
                    Assert.Equal(NativeStatus.Ok, NativeApi.OpenFile(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Auto, out NativeHandle? handle));
                    Assert.NotNull(handle);

                    Span<byte> errorBuffer = stackalloc byte[256];
                    Assert.Equal(NativeStatus.Ok, NativeApi.LastError(errorBuffer, out int length));
                    Assert.Equal(0, length); // this thread never provoked an error - must not see another thread's

                    Assert.Equal(NativeStatus.Ok, NativeApi.Close(handle));
                }
            }));

            return Task.WhenAll(tasks);
        }
    }
}
