using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Reader
{
    public class AsyncEnumerationTests
    {
        [Fact]
        public async Task XlsxAsyncEnumerationNeverReadsSynchronously()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = await BuildAsync<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(s => XlsxWorkbookWriter.Create(s, leaveOpen: true));
            await using var stream = new SyncGuardStream(bytes);
            await using XlsxReader reader = await Excel.FromXlsxAsync(stream, ct: ct);
            stream.AllowSyncReads = false;

            Assert.Equal(["a", "b"], await ReadFirstColumnAsync(reader.GetAsyncEnumerator(ct)));
        }

        [Fact]
        public async Task XlsbAsyncEnumerationNeverReadsSynchronously()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = await BuildAsync<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(s => XlsbWorkbookWriter.Create(s, leaveOpen: true));
            await using var stream = new SyncGuardStream(bytes);
            await using XlsbReader reader = await Excel.FromXlsbAsync(stream, ct: ct);
            stream.AllowSyncReads = false;

            Assert.Equal(["a", "b"], await ReadFirstColumnAsync(reader.GetAsyncEnumerator(ct)));
        }

        private static async Task<List<string>> ReadFirstColumnAsync(IExcelRowEnumerator e)
        {
            List<string> values = [];
            await using (e)
            {
                while (await e.MoveNextAsync())
                {
                    values.Add(e.Current[0].GetString());
                }
            }
            return values;
        }

        private static async Task<byte[]> BuildAsync<TWorkbook, TSheet, TRow>(Func<Stream, TWorkbook> create)
            where TWorkbook : IWorkbookWriter<TSheet>
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            using MemoryStream ms = new();
            await using (TWorkbook wb = create(ms))
            {
                TSheet sheet = wb.AddSheet("S1");
                foreach (string value in (string[])["a", "b"])
                {
                    await using TRow row = await sheet.StartRowAsync(TestContext.Current.CancellationToken);
                    row.Write(value);
                }
                await sheet.EndAsync(TestContext.Current.CancellationToken);
                await wb.EndAsync(TestContext.Current.CancellationToken);
            }
            return ms.ToArray();
        }

        private sealed class SyncGuardStream(byte[] bytes) : Stream
        {
            private readonly MemoryStream _inner = new(bytes, writable: false);

            public bool AllowSyncReads { get; set; } = true;

            public override bool CanRead => true;
            public override bool CanSeek => true;
            public override bool CanWrite => false;
            public override long Length => _inner.Length;
            public override long Position { get => _inner.Position; set => _inner.Position = value; }

            public override int Read(byte[] buffer, int offset, int count)
            {
                return Read(buffer.AsSpan(offset, count));
            }

            public override int Read(Span<byte> buffer)
            {
                return AllowSyncReads ? _inner.Read(buffer) : throw new InvalidOperationException("Synchronous read after the reader was opened.");
            }

            public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
            {
                return _inner.ReadAsync(buffer, cancellationToken);
            }

            public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
            {
                return _inner.ReadAsync(buffer, offset, count, cancellationToken);
            }

            public override long Seek(long offset, SeekOrigin origin)
            {
                return _inner.Seek(offset, origin);
            }

            public override void Flush()
            {
            }

            public override void SetLength(long value)
            {
                throw new NotSupportedException();
            }

            public override void Write(byte[] buffer, int offset, int count)
            {
                throw new NotSupportedException();
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _inner.Dispose();
                }
                base.Dispose(disposing);
            }
        }
    }
}
