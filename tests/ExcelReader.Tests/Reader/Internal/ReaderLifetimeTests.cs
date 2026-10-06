using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Tests.Reader.Internal
{
    public sealed class ReaderLifetimeTests
    {
        [Fact]
        public void Close_Alone_Releases()
        {
            int released = 0;
            ReaderLifetime lifetime = new(() => released++);
            lifetime.Close();
            Assert.Equal(1, released);
        }

        [Fact]
        public void Release_Happens_When_The_Last_Holder_Lets_Go()
        {
            int released = 0;
            ReaderLifetime lifetime = new(() => released++);
            lifetime.Acquire(this);
            lifetime.Acquire(this);

            lifetime.Close();
            Assert.Equal(0, released);
            lifetime.Release();
            Assert.Equal(0, released);
            lifetime.Release();
            Assert.Equal(1, released);
        }

        [Fact]
        public void Close_Twice_Releases_Once()
        {
            int released = 0;
            ReaderLifetime lifetime = new(() => released++);
            lifetime.Close();
            lifetime.Close();
            Assert.Equal(1, released);
        }

        [Fact]
        public void Acquire_After_Close_Throws_Even_While_Holders_Remain()
        {
            ReaderLifetime lifetime = new(() => { });
            lifetime.Acquire(this);
            lifetime.Close();
            ObjectDisposedException ex = Assert.Throws<ObjectDisposedException>(() => lifetime.Acquire(this));
            Assert.Contains(nameof(ReaderLifetimeTests), ex.ObjectName, StringComparison.Ordinal);
        }

        [Fact]
        public void Concurrent_Holders_Release_Exactly_Once()
        {
            int released = 0;
            ReaderLifetime lifetime = new(() => Interlocked.Increment(ref released));
            Parallel.For(0, 64, _ =>
            {
                lifetime.Acquire(this);
                lifetime.Release();
            });
            Assert.Equal(0, released);
            lifetime.Close();
            Assert.Equal(1, released);
        }
    }
}
