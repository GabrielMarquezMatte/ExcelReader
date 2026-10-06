using System.Globalization;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Reader.Xlsx
{
    public sealed class XlsxLifetimeTests
    {
        private static byte[] BuildXlsx(int rows = 300)
        {
            using MemoryStream buffer = new();
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(buffer, leaveOpen: true, new XlsxWriterOptions { UseSharedStrings = true }))
            {
                XlsxSheetWriter sheet = workbook.AddSheet("s");
                for (int r = 0; r < rows; r++)
                {
                    using XlsxRowWriter row = sheet.StartRow();
                    row.Write("name-" + (r % 40).ToString(CultureInfo.InvariantCulture));
                    row.Write(r);
                }
                sheet.End();
                sheet.Dispose();
                workbook.End();
            }
            return buffer.ToArray();
        }

        private static List<string> Drain(XlsxWorkbook.Enumerator e)
        {
            List<string> values = [];
            while (e.MoveNext())
            {
                values.Add(e.Current[0].GetString());
            }
            return values;
        }

        [Fact]
        public void A_Failed_Shared_String_Load_Fails_Every_Later_Enumeration()
        {
            using XlsxWorkbook reader = Excel.FromXlsx(BuildXlsx(), new ExcelReaderOptions { MaxSharedStringBytes = 16 });
            Assert.Throws<ExcelLimitExceededException>(() => reader.GetEnumerator());
            Assert.Throws<ExcelLimitExceededException>(() => reader.GetEnumerator());
        }

        [Fact]
        public async Task A_Failed_Shared_String_Load_Fails_Every_Later_Async_Enumeration()
        {
            using XlsxWorkbook reader = Excel.FromXlsx(
                new MemoryStream(BuildXlsx(), writable: false), leaveOpen: false, new ExcelReaderOptions { MaxSharedStringBytes = 16 });
            for (int attempt = 0; attempt < 2; attempt++)
            {
                await using XlsxWorkbook.Enumerator e = reader.GetAsyncEnumerator(TestContext.Current.CancellationToken);
                await Assert.ThrowsAsync<ExcelLimitExceededException>(async () => await e.MoveNextAsync());
            }
        }

        [Fact]
        public void GetEnumerator_After_Dispose_Throws()
        {
            XlsxWorkbook reader = Excel.FromXlsx(BuildXlsx());
            reader.Dispose();
            Assert.Throws<ObjectDisposedException>(() => reader.GetEnumerator());
            Assert.Throws<ObjectDisposedException>(() => reader.GetAsyncEnumerator(TestContext.Current.CancellationToken));
        }

        [Fact]
        public void An_Enumerator_Outlives_Its_Reader()
        {
            byte[] bytes = BuildXlsx();
            List<string> expected;
            using (XlsxWorkbook reference = Excel.FromXlsx(bytes))
            {
                using XlsxWorkbook.Enumerator all = reference.GetEnumerator();
                expected = Drain(all);
            }

            TrickleStream stream = new(bytes);
            XlsxWorkbook reader = Excel.FromXlsx(stream, leaveOpen: false);
            XlsxWorkbook.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());

            reader.Dispose();
            Assert.True(stream.CanRead);

            List<string> rest = Drain(e);
            Assert.Equal(expected.Skip(1), rest, StringComparer.Ordinal);

            e.Dispose();
            Assert.False(stream.CanRead);
        }

        [Fact]
        public void Disposing_Reader_And_Enumerator_Twice_Is_Harmless()
        {
            TrickleStream stream = new(BuildXlsx(rows: 20));
            XlsxWorkbook reader = Excel.FromXlsx(stream, leaveOpen: false);
            XlsxWorkbook.Enumerator e = reader.GetEnumerator();
            e.Dispose();
            e.Dispose();
            Assert.True(stream.CanRead);
            reader.Dispose();
            reader.Dispose();
            Assert.False(stream.CanRead);
        }

        [Fact]
        public void A_Failed_GetEnumerator_Does_Not_Keep_The_Reader_Alive()
        {
            TrickleStream stream = new(BuildXlsx(rows: 20));
            XlsxWorkbook reader = Excel.FromXlsx(stream, leaveOpen: false, new ExcelReaderOptions { MaxSharedStringBytes = 16 });
            Assert.Throws<ExcelLimitExceededException>(() => reader.GetEnumerator());
            reader.Dispose();
            Assert.False(stream.CanRead);
        }

        [Fact]
        public async Task A_Disposed_Deferred_Enumerator_Does_Not_Reopen_Its_Sheet()
        {
            TrickleStream stream = new(BuildXlsx(rows: 20));
            XlsxWorkbook reader = Excel.FromXlsx(stream, leaveOpen: false);
            XlsxWorkbook.Enumerator e = reader.GetAsyncEnumerator(TestContext.Current.CancellationToken);
            await e.DisposeAsync();
            Task<bool> move = Task.Run(async () => await e.MoveNextAsync(), TestContext.Current.CancellationToken);
            Assert.Same(move, await Task.WhenAny(move, Task.Delay(TimeSpan.FromSeconds(3), TestContext.Current.CancellationToken)));
            Assert.False(await move);
            reader.Dispose();
            Assert.False(stream.CanRead);
        }

        [Fact]
        public async Task A_Cancelled_Shared_String_Load_Is_Undone_And_Loads_Again()
        {
            byte[] bytes = BuildXlsx();
            List<string> expected;
            using (XlsxWorkbook reference = Excel.FromXlsx(bytes))
            {
                using XlsxWorkbook.Enumerator all = reference.GetEnumerator();
                expected = Drain(all);
            }

            ReadOnlySpan<byte> entryName = "xl/sharedStrings.xml"u8;
            int header = bytes.AsSpan().IndexOf(entryName) - 30;
            long bodyStart = header + 30 + entryName.Length + BitConverter.ToUInt16(bytes, header + 28);
            long bodyEnd = bodyStart + BitConverter.ToUInt32(bytes, header + 18);

            using CancellationTokenSource cts = new();
            int reads = 0;
            long cancelledAt = -1;
            TrickleStream stream = new(bytes);
            XlsxWorkbook reader = Excel.FromXlsx(stream, leaveOpen: false);
            stream.OnRead = () =>
            {
                if (++reads == 20)
                {
                    cancelledAt = stream.Position;
                    cts.Cancel();
                }
            };

            await using (XlsxWorkbook.Enumerator cancelled = reader.GetAsyncEnumerator(cts.Token))
            {
                await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await cancelled.MoveNextAsync());
            }
            Assert.InRange(cancelledAt, bodyStart, bodyEnd - 1);

            List<string> actual = [];
            await using (XlsxWorkbook.Enumerator again = reader.GetAsyncEnumerator(TestContext.Current.CancellationToken))
            {
                while (await again.MoveNextAsync())
                {
                    actual.Add(again.Current[0].GetString());
                }
            }
            Assert.Equal(expected, actual, StringComparer.Ordinal);

            reader.Dispose();
            Assert.False(stream.CanRead);
        }
    }
}
