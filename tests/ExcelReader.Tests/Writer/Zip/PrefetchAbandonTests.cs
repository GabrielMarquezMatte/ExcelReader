using System.Diagnostics;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer.Zip
{
    public class PrefetchAbandonTests
    {
        private const int Iterations = 200;
        private const int Rows = 4_000;
        private static readonly string[] Texts = BuildTexts();

        private static string[] BuildTexts()
        {
            uint state = 11;
            var texts = new string[Rows];
            var chars = new char[120];
            for (int i = 0; i < texts.Length; i++)
            {
                for (int c = 0; c < chars.Length; c++)
                {
                    state = (state * 1_664_525) + 1_013_904_223;
                    chars[c] = (char)('0' + (state >> 26));
                }
                texts[i] = new string(chars);
            }
            return texts;
        }

        private static (IDisposable Workbook, Func<bool> Released) StartXlsx(Stream output)
        {
            XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(output, leaveOpen: true, new XlsxWriterOptions { PrefetchWrite = true });
            XlsxSheetWriter sheet = wb.AddSheet("Data");
            for (int i = 0; i < Rows; i++)
            {
                using XlsxRowWriter row = sheet.StartRow();
                row.Write(i);
                row.Write(Texts[i]);
            }
            sheet.StartRow().Write("left open");
            return (wb, () => sheet.ResourcesReleased);
        }

        private static (IDisposable Workbook, Func<bool> Released) StartXlsb(Stream output)
        {
            XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(output, leaveOpen: true, new XlsbWriterOptions { PrefetchWrite = true });
            XlsbSheetWriter sheet = wb.AddSheet("Data");
            for (int i = 0; i < Rows; i++)
            {
                using XlsbRowWriter row = sheet.StartRow();
                row.Write(i);
                row.Write(Texts[i]);
            }
            sheet.StartRow().Write("left open");
            return (wb, () => sheet.ResourcesReleased);
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public async Task DisposingWithARowLeftOpenStopsTheBackgroundWriterFirst(bool xlsb, bool async)
        {
            var outputs = new List<ProbeStream>(Iterations);
            for (int iteration = 0; iteration < Iterations; iteration++)
            {
                var output = new ProbeStream();
                outputs.Add(output);
                (IDisposable workbook, Func<bool> released) = xlsb ? StartXlsb(output) : StartXlsx(output);
                if (async)
                {
                    await Assert.ThrowsAsync<InvalidOperationException>(() => ((IAsyncDisposable)workbook).DisposeAsync().AsTask());
                }
                else
                {
                    Assert.Throws<InvalidOperationException>(workbook.Dispose);
                }
                output.Seal();
                Assert.True(released());
            }
            await Task.Delay(100, TestContext.Current.CancellationToken);
            Assert.All(outputs, static output =>
            {
                Assert.True(output.Writes > 1);
                Assert.Equal(0, output.Overlaps);
                Assert.Equal(0, output.LateWrites);
            });
        }

        /// <summary>A slow forward-only output that counts overlapping writes and writes made after <see cref="Seal"/>.</summary>
        private sealed class ProbeStream : Stream
        {
            private int _active;
            private int _overlaps;
            private int _lateWrites;
            private int _writes;
            private volatile bool _sealed;

            internal int Overlaps => Volatile.Read(ref _overlaps);
            internal int LateWrites => Volatile.Read(ref _lateWrites);
            internal int Writes => Volatile.Read(ref _writes);
            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => throw new NotSupportedException();
            public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }

            internal void Seal()
            {
                _sealed = true;
            }

            public override void Flush() { }

            public override int Read(byte[] buffer, int offset, int count)
            {
                throw new NotSupportedException();
            }

            public override long Seek(long offset, SeekOrigin origin)
            {
                throw new NotSupportedException();
            }

            public override void SetLength(long value)
            {
                throw new NotSupportedException();
            }

            public override void Write(byte[] buffer, int offset, int count)
            {
                Record();
            }

            public override void Write(ReadOnlySpan<byte> buffer)
            {
                Record();
            }

            public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
            {
                Record();
                return ValueTask.CompletedTask;
            }

            private void Record()
            {
                Interlocked.Increment(ref _writes);
                if (_sealed)
                {
                    Interlocked.Increment(ref _lateWrites);
                }
                if (Interlocked.Increment(ref _active) != 1)
                {
                    Interlocked.Increment(ref _overlaps);
                }
                long until = Stopwatch.GetTimestamp() + (Stopwatch.Frequency / 5_000);
                while (Stopwatch.GetTimestamp() < until)
                {
                    Thread.SpinWait(20);
                }
                Interlocked.Decrement(ref _active);
            }
        }
    }
}
