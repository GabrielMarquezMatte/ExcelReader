using System.Text;
using ExcelReader.Native;

namespace ExcelReader.Tests
{
    /// <summary>
    /// Exercises <see cref="NativeHandleTable"/> through <see cref="Exports"/>'s <c>Resolve</c>/
    /// <c>TryFree</c> helpers — the [UnmanagedCallersOnly] entry points themselves cannot be called
    /// from managed code, but these internal helpers behind them can.
    /// </summary>
    /// <remarks>
    /// Handle values used to be raw GCHandle pointers, and GCHandle table slots are recycled — a
    /// value could stop being invalid and silently start naming a different, live workbook after
    /// enough churn. That made "does a double-close ever resolve the wrong object" untestable: the
    /// real hazard could not be reproduced deterministically, and probing near it under the xUnit
    /// test host (itself heavily GCHandle-churning) risked corrupting unrelated tests. Handle values
    /// are now ids from a monotonic counter in <see cref="NativeHandleTable"/> that are never
    /// reissued once retired, so the actual guarantee — a closed id stays invalid forever — is
    /// directly, deterministically testable below.
    /// </remarks>
    public sealed class ExportsHandleSafetyTests
    {
        private static readonly string XlsxFixture = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");

        [Fact]
        public void Resolve_Should_Return_Null_For_A_Zero_Handle()
        {
            Assert.Null(Exports.Resolve(0));
        }

        [Fact]
        public void Register_Then_Resolve_Should_Round_Trip_The_Same_Instance()
        {
            NativeHandle handle = OpenRealHandle();
            nint id = NativeHandleTable.Register(handle);

            Assert.Same(handle, Exports.Resolve(id));

            Assert.True(Exports.TryFree(id, out NativeHandle? freed));
            NativeApi.Close(freed);
        }

        [Fact]
        public void TryFree_Should_Fail_The_Second_Time_On_The_Same_Id()
        {
            NativeHandle handle = OpenRealHandle();
            nint id = NativeHandleTable.Register(handle);

            Assert.True(Exports.TryFree(id, out NativeHandle? target));
            NativeApi.Close(target);

            // The double-close case: a second TryFree on the same id must fail cleanly, never
            // resolve to (or free) some other, unrelated live handle.
            Assert.False(Exports.TryFree(id, out NativeHandle? second));
            Assert.Null(second);
        }

        [Fact]
        public void Resolve_Should_Return_Null_For_A_Retired_Id_Even_After_Further_Registrations()
        {
            NativeHandle handle = OpenRealHandle();
            nint retiredId = NativeHandleTable.Register(handle);
            Assert.True(Exports.TryFree(retiredId, out NativeHandle? target));
            NativeApi.Close(target);

            for (int i = 0; i < 1000; i++)
            {
                NativeHandle other = OpenRealHandle();
                nint otherId = NativeHandleTable.Register(other);

                // The retired id must never be handed back out, and must never resolve again.
                Assert.NotEqual(retiredId, otherId);
                Assert.Null(Exports.Resolve(retiredId));

                Assert.True(Exports.TryFree(otherId, out NativeHandle? freed));
                NativeApi.Close(freed);
            }
        }

        [Fact]
        public void Concurrent_Register_Should_Produce_Distinct_Ids_And_Lose_None()
        {
            const int perThread = 200;
            const int threadCount = 8;
            var ids = new nint[threadCount][];
            var handles = new NativeHandle[threadCount][];

            Parallel.For(0, threadCount, threadIndex =>
            {
                ids[threadIndex] = new nint[perThread];
                handles[threadIndex] = new NativeHandle[perThread];
                for (int i = 0; i < perThread; i++)
                {
                    NativeHandle handle = OpenRealHandle();
                    handles[threadIndex][i] = handle;
                    ids[threadIndex][i] = NativeHandleTable.Register(handle);
                }
            });

            var allIds = ids.SelectMany(x => x).ToArray();
            Assert.Equal(threadCount * perThread, allIds.Distinct().Count());

            for (int threadIndex = 0; threadIndex < threadCount; threadIndex++)
            {
                for (int i = 0; i < perThread; i++)
                {
                    Assert.Same(handles[threadIndex][i], Exports.Resolve(ids[threadIndex][i]));
                    Assert.True(Exports.TryFree(ids[threadIndex][i], out NativeHandle? freed));
                    NativeApi.Close(freed);
                }
            }
        }

        private static NativeHandle OpenRealHandle()
        {
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenFile(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Auto, out NativeHandle? handle));
            Assert.NotNull(handle);
            return handle;
        }
    }
}
