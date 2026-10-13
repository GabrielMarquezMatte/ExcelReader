using ExcelReader.Core.Writer.Internal;
using ExcelReader.Tests.Writer.Zip;

namespace ExcelReader.Tests.Writer
{
    public class WriteOffloadDrainTests
    {
        private sealed class SlowStream : MemoryStream
        {
            public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
            {
                await Task.Delay(15, cancellationToken);
                await base.WriteAsync(buffer, cancellationToken);
            }
        }

        [Fact]
        public async Task DrainWaitsForEverythingWrittenBefore()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var inner = new SlowStream();
            await using var offload = new WriteOffloadStream(inner);
            for (int i = 0; i < 6; i++)
            {
                await offload.WriteAsync(new byte[] { (byte)i, (byte)i }, ct);
            }

            await offload.DrainAsync(ct);

            Assert.Equal([0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5], inner.ToArray());
        }

        [Fact]
        public async Task TheStreamStaysWritableAfterADrain()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var inner = new MemoryStream();
            var offload = new WriteOffloadStream(inner);
            await offload.WriteAsync("ab"u8.ToArray(), ct);
            await offload.DrainAsync(ct);
            await offload.DrainAsync(ct);
            await offload.WriteAsync("cd"u8.ToArray(), ct);
            await offload.FlushAsync(ct);

            Assert.Equal("abcd"u8.ToArray(), inner.ToArray());
            await offload.DisposeAsync();
        }

        [Fact]
        public async Task DrainSurfacesAConsumerFailureWithItsType()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var offload = new WriteOffloadStream(new FailingWriteStream(0));
            await offload.WriteAsync("ab"u8.ToArray(), ct);

            await Assert.ThrowsAsync<IOException>(async () => await offload.DrainAsync(ct));
            await Assert.ThrowsAsync<IOException>(async () => await offload.DisposeAsync());
        }
    }
}
