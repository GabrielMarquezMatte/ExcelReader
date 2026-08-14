using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Native;

namespace ExcelReader.Tests
{
    /// <summary>
    /// Exercises the raw GCHandle handling in <see cref="Exports"/> directly. The
    /// [UnmanagedCallersOnly] entry points themselves cannot be called from managed code, but the
    /// internal <c>Resolve</c>/<c>TryFree</c> helpers behind them can — this is where the stale-handle
    /// bug lived (Exports.cs did GCHandle.FromIntPtr outside any try/catch, which crashes the process
    /// on an invalid handle value instead of returning InvalidHandle).
    /// </summary>
    /// <remarks>
    /// <para>
    /// Deliberately does NOT test a totally arbitrary bit-pattern handle (e.g. 0x12345678): that
    /// reliably access-violates the whole test process rather than throwing a catchable
    /// InvalidOperationException — GCHandle.FromIntPtr only validates cheaply for values that were
    /// at least once a real handle (already-freed/reused), which is also the only realistic case an
    /// ABI caller can produce by using a handle after xl_close. A value fabricated out of thin air is
    /// undefined behavior by the C ABI's own contract (see excelreader.h) and is not defensible from
    /// managed code.
    /// </para>
    /// <para>
    /// Also deliberately does NOT drive <see cref="Exports.TryFree"/> with an already-freed handle
    /// value under the real xUnit process: this test host is itself heavily GCHandle-churning (IPC
    /// marshaling on background threads), so the exact freed slot can legitimately be handed back out
    /// to some unrelated live object before the very next line runs. Reading that slot
    /// (<see cref="Exports.Resolve"/>) is harmless either way, but <c>TryFree</c> would call
    /// <c>GCHandle.Free()</c> on it — freeing a live object that belongs to something else entirely.
    /// That reliably corrupted the process during development of this fix (unrelated fuzz tests
    /// elsewhere in the suite started crashing with "Internal CLR error"). <c>TryFree</c>'s stale-path
    /// is instead covered by symmetry: it wraps GCHandle.FromIntPtr/.Target/.Free() in the exact same
    /// try/catch(InvalidOperationException) shape as <c>Resolve</c> (see Exports.cs), which this file
    /// does verify end to end via the zero-handle and live-handle-round-trip cases below.
    /// </para>
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
        public void Resolve_Should_Return_Null_Instead_Of_Throwing_For_An_Already_Freed_Handle()
        {
            // Read-only (no Free() call), so even in the rare case the process reuses this exact
            // slot for something else before the check runs, this only ever reads a value — it
            // cannot corrupt or free anyone else's object. That is what makes this scenario safe to
            // retry, unlike the TryFree case documented on the class above: a real bug in the catch
            // logic fails every attempt, a slot legitimately reused by unrelated concurrent activity
            // clears on the next one.
            bool sawNull = false;
            for (int attempt = 0; attempt < 25 && !sawNull; attempt++)
            {
                nint stale = GCHandle.ToIntPtr(GCHandle.Alloc(new object()));
                GCHandle.FromIntPtr(stale).Free();

                // Must not throw InvalidOperationException across what would be the ABI boundary.
                sawNull = Exports.Resolve(stale) is null;
            }

            Assert.True(sawNull);
        }

        [Fact]
        public void TryFree_Should_Free_And_Return_The_Target_For_A_Live_Handle()
        {
            // Mirrors what Exports.OpenFile actually does: open a real workbook, then wrap its
            // NativeHandle in a GCHandle by hand so this test controls the raw pointer.
            Assert.Equal(NativeStatus.Ok, NativeApi.OpenFile(Encoding.UTF8.GetBytes(XlsxFixture), NativeFormat.Auto, out NativeHandle? handle));
            Assert.NotNull(handle);
            nint pointer = GCHandle.ToIntPtr(GCHandle.Alloc(handle));

            bool ok = Exports.TryFree(pointer, out NativeHandle? target);

            Assert.True(ok);
            Assert.Same(handle, target);
            NativeApi.Close(target);
        }
    }
}
