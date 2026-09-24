using System.IO.Compression;
using ExcelReader.Core;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer
{
    public class SheetWriterFailureTests
    {
        private const int SpillingRows = 5_000;

        public static TheoryData<string, bool, bool, bool> ZipSheetCases
        {
            get
            {
                TheoryData<string, bool, bool, bool> data = [];
                foreach (string format in (string[])["xlsx", "xlsb"])
                {
                    foreach (bool async in (bool[])[false, true])
                    {
                        foreach (bool leaveOpen in (bool[])[false, true])
                        {
                            data.Add(format, async, leaveOpen, false);
                            data.Add(format, async, leaveOpen, true);
                        }
                    }
                }
                return data;
            }
        }

        public static TheoryData<string, bool, bool> ZipFormatsAsyncAndOwnership => new()
        {
            { "xlsx", false, false },
            { "xlsx", false, true },
            { "xlsx", true, false },
            { "xlsx", true, true },
            { "xlsb", false, false },
            { "xlsb", false, true },
            { "xlsb", true, false },
            { "xlsb", true, true },
        };

        public static TheoryData<string, bool> ZipFormatsAndOwnership => new()
        {
            { "xlsx", false },
            { "xlsx", true },
            { "xlsb", false },
            { "xlsb", true },
        };

        public static TheoryData<bool, bool> AsyncAndOwnership => new()
        {
            { false, false },
            { false, true },
            { true, false },
            { true, true },
        };

        private static IWorkbookWriter<IDisposable> Create(
            string format, Stream stream, bool leaveOpen, bool prefetch = false, CompressionLevel compression = CompressionLevel.Optimal)
        {
            return format switch
            {
                "xlsx" => XlsxWorkbookWriter.Create(stream, leaveOpen, new XlsxWriterOptions { PrefetchWrite = prefetch, Compression = compression }),
                "xlsb" => XlsbWorkbookWriter.Create(stream, leaveOpen, new XlsbWriterOptions { PrefetchWrite = prefetch, Compression = compression }),
                _ => XlsWorkbookWriter.Create(stream, leaveOpen),
            };
        }

        private static void WriteRow(IDisposable sheet, string text)
        {
            switch (sheet)
            {
                case XlsxSheetWriter xlsx:
                    using (XlsxRowWriter row = xlsx.StartRow())
                    {
                        row.Write(text);
                    }
                    break;
                case XlsbSheetWriter xlsb:
                    using (XlsbRowWriter row = xlsb.StartRow())
                    {
                        row.Write(text);
                    }
                    break;
                default:
                    using (XlsRowWriter row = ((XlsSheetWriter)sheet).StartRow())
                    {
                        row.Write(text);
                    }
                    break;
            }
        }

        private static void EndSheet(IDisposable sheet)
        {
            switch (sheet)
            {
                case XlsxSheetWriter xlsx: xlsx.End(); break;
                case XlsbSheetWriter xlsb: xlsb.End(); break;
                default: ((XlsSheetWriter)sheet).End(); break;
            }
        }

        private static ValueTask EndSheetAsync(IDisposable sheet, CancellationToken ct)
        {
            return sheet switch
            {
                XlsxSheetWriter xlsx => xlsx.EndAsync(ct),
                XlsbSheetWriter xlsb => xlsb.EndAsync(ct),
                _ => ((XlsSheetWriter)sheet).EndAsync(ct),
            };
        }

        private static bool Released(IDisposable sheet)
        {
            return sheet switch
            {
                XlsxSheetWriter xlsx => xlsx.ResourcesReleased,
                XlsbSheetWriter xlsb => xlsb.ResourcesReleased,
                _ => ((XlsSheetWriter)sheet).ResourcesReleased,
            };
        }

        private static Task DisposeSheet(IDisposable sheet, bool async)
        {
            if (async)
            {
                return ((IAsyncDisposable)sheet).DisposeAsync().AsTask();
            }
            sheet.Dispose();
            return Task.CompletedTask;
        }

        private static void FailWrites(TrackingStream stream)
        {
            stream.OnWrite = static () => throw new IOException("disk full");
        }

        private static async Task AssertFaultedAndReleased(
            CountingPool pool, IWorkbookWriter<IDisposable> workbook, IDisposable sheet, TrackingStream stream, bool leaveOpen, bool async)
        {
            Assert.True(Released(sheet));
            Assert.Null(await Record.ExceptionAsync(() => DisposeSheet(sheet, async)));
            Assert.Throws<ObjectDisposedException>(() => EndSheet(sheet));
            Assert.Throws<ObjectDisposedException>(() => workbook.AddSheet("S2"));
            Assert.Throws<ObjectDisposedException>(workbook.End);

            Assert.Null(await Record.ExceptionAsync(async () =>
            {
                if (async)
                {
                    await workbook.DisposeAsync();
                }
                else
                {
                    workbook.Dispose();
                }
            }));
            Assert.Equal(!leaveOpen, stream.Disposed);
            pool.AssertEveryBufferReturnedOnce();
        }

        [Theory]
        [MemberData(nameof(ZipSheetCases))]
        public async Task Should_ReleaseTheSheetAndFaultTheWorkbook_When_TheSheetFinalizerFails(
            string format, bool async, bool leaveOpen, bool prefetch)
        {
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = Create(format, stream, leaveOpen, prefetch);
            IDisposable sheet = workbook.AddSheet("S1");
            for (int i = 0; i < (prefetch ? SpillingRows : 1); i++)
            {
                WriteRow(sheet, $"row {i} with some text");
            }
            FailWrites(stream);

            await Assert.ThrowsAsync<IOException>(() => DisposeSheet(sheet, async));

            await AssertFaultedAndReleased(pool, workbook, sheet, stream, leaveOpen, async);
        }

        [Theory]
        [MemberData(nameof(ZipFormatsAsyncAndOwnership))]
        public async Task Should_ReleaseTheSheet_When_ARowFlushFailedBeforeTheFinalizer(string format, bool async, bool leaveOpen)
        {
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = Create(format, stream, leaveOpen, compression: CompressionLevel.NoCompression);
            IDisposable sheet = workbook.AddSheet("S1");
            WriteRow(sheet, "first");
            FailWrites(stream);

            Exception? rowFailure = null;
            for (int i = 0; i < 100_000 && rowFailure is null; i++)
            {
                string text = $"row {i} with some text to fill the buffer";
                rowFailure = async
                    ? await Record.ExceptionAsync(() => WriteRowAsync(sheet, text))
                    : Record.Exception(() => WriteRow(sheet, text));
            }
            Assert.IsType<IOException>(rowFailure);
            Assert.True(Released(sheet));
            Assert.Throws<ObjectDisposedException>(() => WriteRow(sheet, "after the failure"));

            await AssertFaultedAndReleased(pool, workbook, sheet, stream, leaveOpen, async);
        }

        private static async Task WriteRowAsync(IDisposable sheet, string text)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            if (sheet is XlsxSheetWriter xlsx)
            {
                await using XlsxRowWriter row = await xlsx.StartRowAsync(ct);
                row.Write(text);
                return;
            }
            await using XlsbRowWriter xlsbRow = await ((XlsbSheetWriter)sheet).StartRowAsync(ct);
            xlsbRow.Write(text);
        }

        [Theory]
        [MemberData(nameof(ZipFormatsAndOwnership))]
        public async Task Should_ReleaseTheSheetAndFaultTheWorkbook_When_TheSheetFinalizerIsCanceled(string format, bool leaveOpen)
        {
            using var cts = new CancellationTokenSource();
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = Create(format, stream, leaveOpen);
            IDisposable sheet = workbook.AddSheet("S1");
            WriteRow(sheet, "value");
            stream.OnWrite = () =>
            {
                cts.Cancel();
                cts.Token.ThrowIfCancellationRequested();
            };

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => EndSheetAsync(sheet, cts.Token).AsTask());

            await AssertFaultedAndReleased(pool, workbook, sheet, stream, leaveOpen, async: true);
        }

        [Theory]
        [InlineData("xlsx")]
        [InlineData("xlsb")]
        [InlineData("xls")]
        public async Task Should_LeaveTheSheetOpen_When_EndAsyncIsCanceledBeforeItStarts(string format)
        {
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = Create(format, stream, leaveOpen: false);
            IDisposable sheet = workbook.AddSheet("S1");
            WriteRow(sheet, "before");

            await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => EndSheetAsync(sheet, new CancellationToken(canceled: true)).AsTask());

            WriteRow(sheet, "after");
            await workbook.DisposeAsync();

            using IExcelRowReader reader = Excel.Open(stream.ToArray().AsMemory());
            using IExcelRowEnumerator rows = reader.GetEnumerator();
            Assert.True(rows.MoveNext());
            Assert.Equal("before", rows.Current[0].GetString());
            Assert.True(rows.MoveNext());
            Assert.Equal("after", rows.Current[0].GetString());
            Assert.True(stream.Disposed);
            pool.AssertEveryBufferReturnedOnce();
        }

        [Theory]
        [MemberData(nameof(AsyncAndOwnership))]
        public async Task Should_SurfaceXlsIoAtTheWorkbookAndReleaseSheetBuffers_When_WritesFail(bool async, bool leaveOpen)
        {
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = Create("xls", stream, leaveOpen);
            IDisposable sheet = workbook.AddSheet("S1");
            WriteRow(sheet, "value");
            FailWrites(stream);

            await DisposeSheet(sheet, async);
            Assert.False(Released(sheet));

            if (async)
            {
                await Assert.ThrowsAsync<IOException>(() => workbook.DisposeAsync().AsTask());
            }
            else
            {
                Assert.Throws<IOException>(workbook.Dispose);
            }

            Assert.True(Released(sheet));
            Assert.Equal(leaveOpen ? 0 : 1, stream.DisposeCount);
            pool.AssertEveryBufferReturnedOnce();
        }

        public static TheoryData<string, bool, bool> XlsxFinalizerStepFailures
        {
            get
            {
                TheoryData<string, bool, bool> data = [];
                foreach (string step in (string[])["write", "flush", "dispose"])
                {
                    foreach (bool async in (bool[])[false, true])
                    {
                        data.Add(step, async, false);
                        data.Add(step, async, true);
                    }
                }
                return data;
            }
        }

        [Theory]
        [MemberData(nameof(XlsxFinalizerStepFailures))]
        public async Task Should_ReleaseTheXlsxSheet_When_AnyFinalizerStepFails(string step, bool async, bool leaveOpen)
        {
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = Create("xlsx", stream, leaveOpen, compression: CompressionLevel.NoCompression);
            IDisposable sheet = workbook.AddSheet("S1");
            WriteRow(sheet, "value");
            switch (step)
            {
                case "write":
                    FailWrites(stream);
                    break;
                case "flush":
                    stream.OnFlush = static () => throw new IOException("flush failed");
                    break;
                default:
                    stream.OnFlush = () => FailWrites(stream);
                    break;
            }

            await Assert.ThrowsAsync<IOException>(() => DisposeSheet(sheet, async));

            await AssertFaultedAndReleased(pool, workbook, sheet, stream, leaveOpen, async);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task Should_FaultOnce_When_TheSheetStartFailsInsideItsFinalizer(bool async)
        {
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = Create("xlsx", stream, leaveOpen: false, compression: CompressionLevel.NoCompression);
            IDisposable sheet = workbook.AddSheet("NeverStarted");
            FailWrites(stream);

            await Assert.ThrowsAsync<IOException>(() => DisposeSheet(sheet, async));

            await AssertFaultedAndReleased(pool, workbook, sheet, stream, leaveOpen: false, async);
        }

        [Theory]
        [InlineData("xlsx")]
        [InlineData("xlsb")]
        public void Should_StayFaulted_When_TheOwnerIsNotifiedTwice(string format)
        {
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = Create(format, stream, leaveOpen: false);
            workbook.AddSheet("S1").Dispose();

            switch (workbook)
            {
                case XlsxWorkbookWriter xlsx:
                    xlsx.NotifySheetEnded(faulted: true);
                    xlsx.NotifySheetEnded(faulted: true);
                    break;
                default:
                    ((XlsbWorkbookWriter)workbook).NotifySheetEnded(faulted: true);
                    ((XlsbWorkbookWriter)workbook).NotifySheetEnded(faulted: true);
                    break;
            }

            Assert.Throws<ObjectDisposedException>(workbook.End);
            Assert.Throws<ObjectDisposedException>(() => workbook.AddSheet("S2"));
            workbook.Dispose();
            workbook.Dispose();
            Assert.Equal(1, stream.DisposeCount);
            pool.AssertEveryBufferReturnedOnce();
        }

        [Theory]
        [MemberData(nameof(ZipFormatsAsyncAndOwnership))]
        public async Task Should_CloseTheArchiveExactlyOnce_When_ASheetFaultsOutsideTheWorkbookEnd(string format, bool async, bool leaveOpen)
        {
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = Create(format, stream, leaveOpen);
            IDisposable sheet = workbook.AddSheet("S1");
            WriteRow(sheet, "value");
            stream.OnWrite = TrackingStream.FailOnce(new IOException("disk full"));

            await Assert.ThrowsAsync<IOException>(() => DisposeSheet(sheet, async));
            await Assert.ThrowsAsync<ObjectDisposedException>(() => EndSheetAsync(sheet, CancellationToken.None).AsTask());

            if (async)
            {
                await workbook.DisposeAsync();
                await workbook.DisposeAsync();
            }
            else
            {
                workbook.Dispose();
                workbook.Dispose();
            }

            ReadOnlySpan<byte> endOfCentralDirectory = [0x50, 0x4B, 0x05, 0x06];
            byte[] written = stream.ToArray();
            int first = written.AsSpan().IndexOf(endOfCentralDirectory);
            Assert.True(first >= 0);
            Assert.True(written.AsSpan(first + 4).IndexOf(endOfCentralDirectory) < 0);
            Assert.Equal(leaveOpen ? 0 : 1, stream.DisposeCount);
            pool.AssertEveryBufferReturnedOnce();
        }

        [Fact]
        public void Should_ReleaseXlsSheetBuffers_When_TheWorkbookRejectsAnAllHiddenWorkbook()
        {
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = Create("xls", stream, leaveOpen: false);
            IDisposable sheet = workbook.AddSheet("S1", ExcelSheetVisibility.Hidden);
            WriteRow(sheet, "value");
            sheet.Dispose();

            Assert.Throws<InvalidOperationException>(workbook.Dispose);

            Assert.True(Released(sheet));
            Assert.Equal(1, stream.DisposeCount);
            pool.AssertEveryBufferReturnedOnce();
        }

        [Fact]
        public void Should_KeepTheXlsSheetUsable_When_EndedWhileAContinuationRowIsOpen()
        {
            using var pool = CountingPool.Install();
            var stream = new TrackingStream();
            using (XlsWorkbookWriter workbook = XlsWorkbookWriter.Create(stream, leaveOpen: true))
            {
                XlsSheetWriter sheet = workbook.AddSheet("Big");
                for (int i = 0; i < 65_536; i++)
                {
                    using XlsRowWriter row = sheet.StartRow();
                    row.Write(i);
                }
                XlsRowWriter spilled = sheet.StartRow();
                spilled.Write(65_536);

                Assert.Throws<InvalidOperationException>(sheet.End);

                spilled.Dispose();
                sheet.End();
            }

            using IExcelRowReader reader = Excel.Open(stream.ToArray().AsMemory());
            Assert.Equal(2, reader.SheetCount);
            pool.AssertEveryBufferReturnedOnce();
        }
    }
}
