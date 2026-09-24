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
            IWorkbookWriter<IDisposable> workbook, IDisposable sheet, TrackingStream stream, bool leaveOpen, bool async)
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
        }

        [Theory]
        [MemberData(nameof(ZipSheetCases))]
        public async Task Should_ReleaseTheSheetAndFaultTheWorkbook_When_TheSheetFinalizerFails(
            string format, bool async, bool leaveOpen, bool prefetch)
        {
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = Create(format, stream, leaveOpen, prefetch);
            IDisposable sheet = workbook.AddSheet("S1");
            for (int i = 0; i < (prefetch ? SpillingRows : 1); i++)
            {
                WriteRow(sheet, $"row {i} with some text");
            }
            FailWrites(stream);

            await Assert.ThrowsAsync<IOException>(() => DisposeSheet(sheet, async));

            await AssertFaultedAndReleased(workbook, sheet, stream, leaveOpen, async);
        }

        [Theory]
        [MemberData(nameof(ZipFormatsAsyncAndOwnership))]
        public async Task Should_ReleaseTheSheet_When_ARowFlushFailedBeforeTheFinalizer(string format, bool async, bool leaveOpen)
        {
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

            await AssertFaultedAndReleased(workbook, sheet, stream, leaveOpen, async);
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

            await AssertFaultedAndReleased(workbook, sheet, stream, leaveOpen, async: true);
        }

        [Theory]
        [InlineData("xlsx")]
        [InlineData("xlsb")]
        [InlineData("xls")]
        public async Task Should_LeaveTheSheetOpen_When_EndAsyncIsCanceledBeforeItStarts(string format)
        {
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
        }

        [Theory]
        [MemberData(nameof(AsyncAndOwnership))]
        public async Task Should_SurfaceXlsIoAtTheWorkbookAndReleaseSheetBuffers_When_WritesFail(bool async, bool leaveOpen)
        {
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
            Assert.Equal(!leaveOpen, stream.Disposed);
        }

        [Fact]
        public void Should_ReleaseXlsSheetBuffers_When_TheWorkbookRejectsAnAllHiddenWorkbook()
        {
            var stream = new TrackingStream();
            IWorkbookWriter<IDisposable> workbook = Create("xls", stream, leaveOpen: false);
            IDisposable sheet = workbook.AddSheet("S1", ExcelSheetVisibility.Hidden);
            WriteRow(sheet, "value");
            sheet.Dispose();

            Assert.Throws<InvalidOperationException>(workbook.Dispose);

            Assert.True(Released(sheet));
            Assert.True(stream.Disposed);
        }

        [Fact]
        public void Should_KeepTheXlsSheetUsable_When_EndedWhileAContinuationRowIsOpen()
        {
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
        }
    }
}
