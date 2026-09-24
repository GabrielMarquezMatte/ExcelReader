using ExcelReader.Core;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer
{
    public class WorkbookWriterFailureTests
    {
        public static TheoryData<string, bool> FormatsAndOwnership => new()
        {
            { "xlsx", false },
            { "xlsx", true },
            { "xlsb", false },
            { "xlsb", true },
            { "xls", false },
            { "xls", true },
        };

        private static IWorkbookWriter<IDisposable> Create(string format, Stream stream, bool leaveOpen)
        {
            return format switch
            {
                "xlsx" => XlsxWorkbookWriter.Create(stream, leaveOpen),
                "xlsb" => XlsbWorkbookWriter.Create(stream, leaveOpen),
                _ => XlsWorkbookWriter.Create(stream, leaveOpen),
            };
        }

        private static IWorkbookWriter<IDisposable> CreateWithEndedSheet(
            string format, TrackingStream stream, bool leaveOpen, ExcelSheetVisibility visibility = ExcelSheetVisibility.Visible)
        {
            IWorkbookWriter<IDisposable> workbook = Create(format, stream, leaveOpen);
            workbook.AddSheet("S1", visibility).Dispose();
            return workbook;
        }

        private static void FailWrites(TrackingStream stream)
        {
            stream.OnWrite = static () => throw new IOException("disk full");
        }

        [Theory]
        [MemberData(nameof(FormatsAndOwnership))]
        public void Should_ReleaseOwnedStream_When_DisposeFailsToFinalize(string format, bool leaveOpen)
        {
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = CreateWithEndedSheet(format, stream, leaveOpen);
            FailWrites(stream);

            Assert.Throws<IOException>(workbook.Dispose);

            Assert.Equal(!leaveOpen, stream.Disposed);
            Assert.Null(Record.Exception(workbook.Dispose));
        }

        [Theory]
        [MemberData(nameof(FormatsAndOwnership))]
        public async Task Should_ReleaseOwnedStream_When_DisposeAsyncFailsToFinalize(string format, bool leaveOpen)
        {
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = CreateWithEndedSheet(format, stream, leaveOpen);
            FailWrites(stream);

            await Assert.ThrowsAsync<IOException>(() => workbook.DisposeAsync().AsTask());

            Assert.Equal(!leaveOpen, stream.Disposed);
            Assert.Null(await Record.ExceptionAsync(() => workbook.DisposeAsync().AsTask()));
        }

        [Theory]
        [InlineData("xlsx", false)]
        [InlineData("xlsx", true)]
        [InlineData("xlsb", false)]
        [InlineData("xlsb", true)]
        [InlineData("xls", false)]
        [InlineData("xls", true)]
        public async Task Should_ReportTheFinalizationFailure_When_ClosingTheOwnedStreamAlsoFails(string format, bool async)
        {
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = CreateWithEndedSheet(format, stream, leaveOpen: false);
            FailWrites(stream);
            stream.OnDispose = static () => throw new InvalidOperationException("close failed");

            if (async)
            {
                await Assert.ThrowsAsync<IOException>(() => workbook.DisposeAsync().AsTask());
            }
            else
            {
                Assert.Throws<IOException>(workbook.Dispose);
            }

            Assert.True(stream.Disposed);
        }

        [Theory]
        [MemberData(nameof(FormatsAndOwnership))]
        public void Should_ReleaseOwnedStream_When_DisposeRejectsAnAllHiddenWorkbook(string format, bool leaveOpen)
        {
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = CreateWithEndedSheet(format, stream, leaveOpen, ExcelSheetVisibility.Hidden);

            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(workbook.Dispose);

            Assert.Contains("at least one visible sheet", ex.Message, StringComparison.Ordinal);
            Assert.Equal(!leaveOpen, stream.Disposed);
        }

        [Theory]
        [MemberData(nameof(FormatsAndOwnership))]
        public void Should_StayEnded_When_EndFails(string format, bool leaveOpen)
        {
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = CreateWithEndedSheet(format, stream, leaveOpen);
            FailWrites(stream);

            Assert.Throws<IOException>(workbook.End);
            Assert.Throws<ObjectDisposedException>(workbook.End);
            Assert.Throws<ObjectDisposedException>(() => workbook.AddSheet("S2"));

            stream.OnWrite = null;
            Assert.Null(Record.Exception(workbook.Dispose));
            Assert.Equal(!leaveOpen, stream.Disposed);
        }

        [Theory]
        [MemberData(nameof(FormatsAndOwnership))]
        public async Task Should_StayEnded_When_EndAsyncFails(string format, bool leaveOpen)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = CreateWithEndedSheet(format, stream, leaveOpen);
            FailWrites(stream);

            await Assert.ThrowsAsync<IOException>(() => workbook.EndAsync(ct).AsTask());
            await Assert.ThrowsAsync<ObjectDisposedException>(() => workbook.EndAsync(ct).AsTask());

            stream.OnWrite = null;
            Assert.Null(await Record.ExceptionAsync(() => workbook.DisposeAsync().AsTask()));
            Assert.Equal(!leaveOpen, stream.Disposed);
        }

        [Theory]
        [MemberData(nameof(FormatsAndOwnership))]
        public async Task Should_StayEnded_When_EndAsyncIsCanceledMidway(string format, bool leaveOpen)
        {
            using var cts = new CancellationTokenSource();
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = CreateWithEndedSheet(format, stream, leaveOpen);
            stream.OnWrite = cts.Cancel;

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => workbook.EndAsync(cts.Token).AsTask());
            await Assert.ThrowsAsync<ObjectDisposedException>(() => workbook.EndAsync(CancellationToken.None).AsTask());

            stream.OnWrite = null;
            Assert.Null(await Record.ExceptionAsync(() => workbook.DisposeAsync().AsTask()));
            Assert.Equal(!leaveOpen, stream.Disposed);
        }

        [Theory]
        [MemberData(nameof(FormatsAndOwnership))]
        public async Task Should_StayOpen_When_EndAsyncIsCanceledBeforeItStarts(string format, bool leaveOpen)
        {
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = CreateWithEndedSheet(format, stream, leaveOpen);

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => workbook.EndAsync(new CancellationToken(canceled: true)).AsTask());
            await workbook.DisposeAsync();

            using IExcelRowReader reader = Excel.Open(stream.ToArray().AsMemory());
            Assert.Equal("S1", reader.SheetNameAt(0));
            Assert.Equal(!leaveOpen, stream.Disposed);
        }

        [Theory]
        [InlineData("xlsx", "End")]
        [InlineData("xlsx", "EndAsync")]
        [InlineData("xlsx", "Dispose")]
        [InlineData("xlsx", "DisposeAsync")]
        [InlineData("xlsb", "End")]
        [InlineData("xlsb", "EndAsync")]
        [InlineData("xlsb", "Dispose")]
        [InlineData("xlsb", "DisposeAsync")]
        [InlineData("xls", "End")]
        [InlineData("xls", "EndAsync")]
        [InlineData("xls", "Dispose")]
        [InlineData("xls", "DisposeAsync")]
        public async Task Should_WriteTheActiveSheet_When_ClosedBeforeTheSheetStarted(string format, string close)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = Create(format, stream, leaveOpen: true);
            _ = workbook.AddSheet("Only");

            switch (close)
            {
                case "End": workbook.End(); break;
                case "EndAsync": await workbook.EndAsync(ct); break;
                case "Dispose": workbook.Dispose(); break;
                default: await workbook.DisposeAsync(); break;
            }
            await workbook.DisposeAsync();

            using IExcelRowReader reader = Excel.Open(stream.ToArray().AsMemory());
            Assert.Equal(1, reader.SheetCount);
            Assert.Equal("Only", reader.SheetNameAt(0));
        }
    }
}
