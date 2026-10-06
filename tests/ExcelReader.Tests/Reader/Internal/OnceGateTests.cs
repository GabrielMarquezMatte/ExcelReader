using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Tests.Reader.Internal
{
    public sealed class OnceGateTests
    {
        private sealed class Probe
        {
            internal int Loads;
            internal int Resets;
            internal Exception? Fail;
        }

        private static void Load(Probe probe)
        {
            Interlocked.Increment(ref probe.Loads);
            if (probe.Fail is not null)
            {
                throw probe.Fail;
            }
        }

        private static void Reset(Probe probe)
        {
            Interlocked.Increment(ref probe.Resets);
        }

        [Fact]
        public void Load_Runs_Once_For_Many_Concurrent_Callers()
        {
            using OnceGate gate = new();
            Probe probe = new();
            Parallel.For(0, 64, _ => gate.Run(probe, Load, Reset));
            Assert.Equal(1, probe.Loads);
            Assert.True(gate.IsDone);
        }

        [Fact]
        public void A_Failure_Is_Rethrown_Without_Loading_Again()
        {
            using OnceGate gate = new();
            Probe probe = new() { Fail = new InvalidDataException("bad table") };

            InvalidDataException first = Assert.Throws<InvalidDataException>(() => gate.Run(probe, Load, Reset));
            probe.Fail = null;
            InvalidDataException second = Assert.Throws<InvalidDataException>(() => gate.Run(probe, Load, Reset));

            Assert.Same(first, second);
            Assert.Equal(1, probe.Loads);
            Assert.Equal(0, probe.Resets);
            Assert.False(gate.IsDone);
        }

        [Fact]
        public void Cancellation_Resets_And_Lets_The_Next_Caller_Load()
        {
            using OnceGate gate = new();
            Probe probe = new() { Fail = new OperationCanceledException() };

            Assert.Throws<OperationCanceledException>(() => gate.Run(probe, Load, Reset));
            Assert.Equal(1, probe.Resets);
            Assert.False(gate.IsDone);

            probe.Fail = null;
            gate.Run(probe, Load, Reset);
            Assert.Equal(2, probe.Loads);
            Assert.True(gate.IsDone);
        }

        [Fact]
        public async Task RunAsync_Follows_The_Same_Rules()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using OnceGate gate = new();
            Probe probe = new() { Fail = new OperationCanceledException() };

            static ValueTask LoadAsync(Probe p, CancellationToken token)
            {
                Load(p);
                return ValueTask.CompletedTask;
            }

            await Assert.ThrowsAsync<OperationCanceledException>(async () => await gate.RunAsync(probe, LoadAsync, Reset, ct));
            Assert.Equal(1, probe.Resets);

            probe.Fail = new InvalidDataException("bad table");
            await Assert.ThrowsAsync<InvalidDataException>(async () => await gate.RunAsync(probe, LoadAsync, Reset, ct));
            probe.Fail = null;
            await Assert.ThrowsAsync<InvalidDataException>(async () => await gate.RunAsync(probe, LoadAsync, Reset, ct));
            Assert.Equal(2, probe.Loads);
        }

        [Fact]
        public async Task A_Sync_Caller_And_An_Async_Caller_Share_One_Load()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using OnceGate gate = new();
            Probe probe = new();

            static ValueTask LoadAsync(Probe p, CancellationToken token)
            {
                Load(p);
                return ValueTask.CompletedTask;
            }

            await gate.RunAsync(probe, LoadAsync, Reset, ct);
            gate.Run(probe, Load, Reset);
            Assert.Equal(1, probe.Loads);
        }
    }
}
