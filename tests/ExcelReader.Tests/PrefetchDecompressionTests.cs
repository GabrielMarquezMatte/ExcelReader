using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    public class PrefetchDecompressionTests
    {
        private static readonly TimeSpan HangGuard = TimeSpan.FromSeconds(10);

        public static IEnumerable<object[]> Fixtures
        {
            get
            {
                yield return [new PrefetchFixture("xlsx", BuildLargeXlsx, OpenXlsx, OpenXlsxAsync)];
                yield return [new PrefetchFixture("xlsb", BuildLargeXlsb, OpenXlsb, OpenXlsbAsync)];
            }
        }


        [Theory]
        [MemberData(nameof(Fixtures))]
        public void SyncReadIsIdenticalWithPrefetchOnAndOff(PrefetchFixture fixture)
        {
            byte[] bytes = fixture.Build();
            var off = new ExcelReaderOptions { PrefetchDecompression = false };
            var on = new ExcelReaderOptions { PrefetchDecompression = true };

            List<CellSnapshot> serial = ReadSync(bytes, fixture.OpenSync, off);
            List<CellSnapshot> prefetched = ReadSync(bytes, fixture.OpenSync, on);

            Assert.NotEmpty(serial);
            Assert.Equal(serial, prefetched);
        }

        [Theory]
        [MemberData(nameof(Fixtures))]
        public async Task AsyncReadIsIdenticalWithPrefetchOnAndOff(PrefetchFixture fixture)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = fixture.Build();
            var off = new ExcelReaderOptions { PrefetchDecompression = false };
            var on = new ExcelReaderOptions { PrefetchDecompression = true };

            List<CellSnapshot> serial = await ReadAsync(bytes, fixture.OpenAsync, off, ct);
            List<CellSnapshot> prefetched = await ReadAsync(bytes, fixture.OpenAsync, on, ct);

            Assert.NotEmpty(serial);
            Assert.Equal(serial, prefetched);
        }


        [SuppressMessage("Reliability", "CA2025:Ensure task instances are complete before disposing them",
            Justification = "The disposables live inside the Task.Run delegate and are disposed there before the delegate completes; the returned Task is fully awaited by AssertCompletesWithinGuardAsync.")]
        [Theory]
        [MemberData(nameof(Fixtures))]
        public Task DisposingAfterOneRowCompletesPromptlySync(PrefetchFixture fixture)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = fixture.Build();
            var options = new ExcelReaderOptions { PrefetchDecompression = true };

            Task work = Task.Run(() =>
            {
                using MemoryStream stream = new(bytes, writable: false);
                using IExcelRowReader reader = fixture.OpenSync(stream, options);
                using IExcelRowEnumerator e = reader.GetEnumerator();
                Assert.True(e.MoveNext());
            }, ct);

            return AssertCompletesWithinGuardAsync(work, ct);
        }

        [SuppressMessage("Reliability", "CA2025:Ensure task instances are complete before disposing them",
            Justification = "The disposables live inside the Task.Run delegate and are disposed there before the delegate completes; the returned Task is fully awaited by AssertCompletesWithinGuardAsync.")]
        [Theory]
        [MemberData(nameof(Fixtures))]
        public Task DisposingAfterOneRowCompletesPromptlyAsync(PrefetchFixture fixture)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = fixture.Build();
            var options = new ExcelReaderOptions { PrefetchDecompression = true };

            Task work = Task.Run(async () =>
            {
                await using MemoryStream stream = new(bytes, writable: false);
                await using IExcelRowReader reader = await fixture.OpenAsync(stream, options, ct);
                await using IExcelRowEnumerator e = await reader.GetAsyncEnumeratorAsync(ct);
                Assert.True(await e.MoveNextAsync());
            }, ct);

            return AssertCompletesWithinGuardAsync(work, ct);
        }

        private static async Task AssertCompletesWithinGuardAsync(Task work, CancellationToken ct)
        {
            using CancellationTokenSource delayCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            Task delay = Task.Delay(HangGuard, delayCts.Token);
            Task completed = await Task.WhenAny(work, delay);
            await delayCts.CancelAsync();
            Assert.Same(work, completed);
            await work;
        }


        [Fact]
        public void CorruptEntryThrowsSameExceptionTypeSyncWithPrefetchOnAndOff()
        {
            byte[] bytes = BuildLargeXlsx();
            CorruptEntryCompressedData(bytes, "xl/worksheets/sheet1.xml");

            Type serialType = CaptureSyncExceptionType(bytes, prefetch: false);
            Type prefetchType = CaptureSyncExceptionType(bytes, prefetch: true);

            Assert.Equal(typeof(InvalidDataException), serialType);
            Assert.Equal(serialType, prefetchType);
        }

        [Fact]
        public async Task CorruptEntryThrowsSameExceptionTypeAsyncWithPrefetchOnAndOff()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] bytes = BuildLargeXlsx();
            CorruptEntryCompressedData(bytes, "xl/worksheets/sheet1.xml");

            Type serialType = await CaptureAsyncExceptionTypeAsync(bytes, prefetch: false, ct);
            Type prefetchType = await CaptureAsyncExceptionTypeAsync(bytes, prefetch: true, ct);

            Assert.Equal(typeof(InvalidDataException), serialType);
            Assert.Equal(serialType, prefetchType);
        }

        private static Type CaptureSyncExceptionType(byte[] bytes, bool prefetch)
        {
            var options = new ExcelReaderOptions { PrefetchDecompression = prefetch };
            using MemoryStream stream = new(bytes, writable: false);
            using IExcelRowReader reader = Excel.FromXlsx(stream, options: options);
            Exception ex = Assert.Throws<InvalidDataException>(() =>
            {
                using IExcelRowEnumerator e = reader.GetEnumerator();
                DrainRows(e);
            });
            return ex.GetType();
        }

        private static async Task<Type> CaptureAsyncExceptionTypeAsync(byte[] bytes, bool prefetch, CancellationToken ct)
        {
            var options = new ExcelReaderOptions { PrefetchDecompression = prefetch };
            await using MemoryStream stream = new(bytes, writable: false);
            await using XlsxReader reader = await Excel.FromXlsxAsync(stream, options: options, ct: ct);
            Exception ex = await Assert.ThrowsAsync<InvalidDataException>(async () =>
            {
                await using XlsxReader.Enumerator e = await reader.GetAsyncEnumeratorAsync(ct);
                await DrainRowsAsync(e);
            });
            return ex.GetType();
        }

        private static void DrainRows(IExcelRowEnumerator e)
        {
            while (e.MoveNext())
            {
                for (int c = 0; c < e.Current.ColumnCount; c++)
                {
                    _ = e.Current[c].GetString();
                }
            }
        }

        private static async Task DrainRowsAsync(XlsxReader.Enumerator e)
        {
            while (await e.MoveNextAsync())
            {
                for (int c = 0; c < e.Current.ColumnCount; c++)
                {
                    _ = e.Current[c].GetString();
                }
            }
        }


        [Fact]
        public void PrefetchOnStillTripsTotalDecompressedLimitSync()
        {
            string value = new('A', 256 * 1024);
            using MemoryStream ms = WorkbookBuilder.Build(
                $"""<row r="1"><c r="A1" t="inlineStr"><is><t>{value}</t></is></c></row>""");
            var options = new ExcelReaderOptions
            {
                MaxTotalDecompressedBytes = 16 * 1024,
                PrefetchDecompression = true,
            };

            using XlsxReader reader = Excel.FromXlsx(ms, options: options);
            ExcelLimitExceededException ex = Assert.Throws<ExcelLimitExceededException>(() =>
            {
                using XlsxReader.Enumerator e = reader.GetEnumerator();
                DrainRows(e);
            });
            Assert.Equal(nameof(ExcelReaderOptions.MaxTotalDecompressedBytes), ex.LimitName);
        }

        [Fact]
        public async Task PrefetchOnStillTripsTotalDecompressedLimitAsync()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            string value = new('A', 256 * 1024);
            await using MemoryStream ms = WorkbookBuilder.Build(
                $"""<row r="1"><c r="A1" t="inlineStr"><is><t>{value}</t></is></c></row>""");
            var options = new ExcelReaderOptions
            {
                MaxTotalDecompressedBytes = 16 * 1024,
                PrefetchDecompression = true,
            };

            await using XlsxReader reader = await Excel.FromXlsxAsync(ms, options: options, ct: ct);
            ExcelLimitExceededException ex = await Assert.ThrowsAsync<ExcelLimitExceededException>(async () =>
            {
                await using XlsxReader.Enumerator e = await reader.GetAsyncEnumeratorAsync(ct);
                await DrainRowsAsync(e);
            });
            Assert.Equal(nameof(ExcelReaderOptions.MaxTotalDecompressedBytes), ex.LimitName);
        }


        [Fact]
        public async Task AlreadyCancelledTokenThrowsWithPrefetchOn()
        {
            CancellationToken openCt = TestContext.Current.CancellationToken;
            byte[] bytes = BuildLargeXlsx();
            var options = new ExcelReaderOptions { PrefetchDecompression = true };
            using CancellationTokenSource cts = new();
            cts.Cancel();

            await using MemoryStream stream = new(bytes, writable: false);
            await using XlsxReader reader = await Excel.FromXlsxAsync(stream, options: options, ct: openCt);

            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
                await reader.GetAsyncEnumeratorAsync(cts.Token));
        }

        [Fact]
        public async Task CancellingMidEnumerationStopsPromptlyWithPrefetchOn()
        {
            CancellationToken openCt = TestContext.Current.CancellationToken;
            byte[] bytes = BuildLargeXlsx();
            var options = new ExcelReaderOptions { PrefetchDecompression = true };
            using CancellationTokenSource cts = new();

            await using MemoryStream stream = new(bytes, writable: false);
            await using XlsxReader reader = await Excel.FromXlsxAsync(stream, options: options, ct: openCt);
            await using XlsxReader.Enumerator e = await reader.GetAsyncEnumeratorAsync(cts.Token);

            Assert.True(await e.MoveNextAsync());
            cts.Cancel();

            Task moveNext = Task.Run(async () => await e.MoveNextAsync(), openCt);
            await AssertCompletesWithinGuardAsync(moveNext, openCt);
        }


        [Fact]
        public void WrapSkipsPrefetchForAnEntrySmallerThanTheThreshold()
        {
            using var inner = new MemoryStream(new byte[16]);
            var options = new ExcelReaderOptions { PrefetchDecompression = true };
            using LimitedReadStream wrapped = WorkbookLookups.Wrap(
                inner, new DecompressedByteCounter(0), options, "", 0, uncompressedSize: 16);

            Assert.False(GetInnerStream(wrapped) is PrefetchStream);
        }

        [Fact]
        public void WrapUsesPrefetchForAnEntryAtOrAboveTheThresholdWhenEnabled()
        {
            using var inner = new MemoryStream(new byte[16]);
            var options = new ExcelReaderOptions { PrefetchDecompression = true };
            using LimitedReadStream wrapped = WorkbookLookups.Wrap(
                inner, new DecompressedByteCounter(0), options, "", 0, uncompressedSize: 256 * 1024);

            Assert.True(GetInnerStream(wrapped) is PrefetchStream);
        }

        [Fact]
        public void WrapNeverUsesPrefetchWhenDisabledRegardlessOfSize()
        {
            using var inner = new MemoryStream(new byte[16]);
            var options = new ExcelReaderOptions { PrefetchDecompression = false };
            using LimitedReadStream wrapped = WorkbookLookups.Wrap(
                inner, new DecompressedByteCounter(0), options, "", 0, uncompressedSize: 256 * 1024);

            Assert.False(GetInnerStream(wrapped) is PrefetchStream);
        }

        private static Stream GetInnerStream(LimitedReadStream stream)
        {
            var field = typeof(LimitedReadStream).GetField("_inner", BindingFlags.NonPublic | BindingFlags.Instance)!;
            return (Stream)field.GetValue(stream)!;
        }


        private static List<CellSnapshot> ReadSync(
            byte[] bytes,
            Func<Stream, ExcelReaderOptions, IExcelRowReader> open,
            ExcelReaderOptions options)
        {
            using MemoryStream stream = new(bytes, writable: false);
            using IExcelRowReader reader = open(stream, options);
            using IExcelRowEnumerator e = reader.GetEnumerator();
            List<CellSnapshot> cells = [];
            int rowIndex = 0;
            while (e.MoveNext())
            {
                AddRow(cells, rowIndex++, e.Current);
            }
            return cells;
        }

        private static async Task<List<CellSnapshot>> ReadAsync(
            byte[] bytes,
            Func<Stream, ExcelReaderOptions, CancellationToken, ValueTask<IExcelRowReader>> open,
            ExcelReaderOptions options,
            CancellationToken ct)
        {
            await using MemoryStream stream = new(bytes, writable: false);
            await using IExcelRowReader reader = await open(stream, options, ct);
            await using IExcelRowEnumerator e = await reader.GetAsyncEnumeratorAsync(ct);
            List<CellSnapshot> cells = [];
            int rowIndex = 0;
            while (await e.MoveNextAsync())
            {
                AddRow(cells, rowIndex++, e.Current);
            }
            return cells;
        }

        private static void AddRow(List<CellSnapshot> cells, int rowIndex, Row row)
        {
            cells.Add(CellSnapshot.RowMarker(rowIndex, row.ColumnCount));
            for (int column = 0; column < row.ColumnCount; column++)
            {
                Cell cell = row[column];
                bool hasDouble = cell.TryGetDouble(out double value);
                cells.Add(new CellSnapshot(
                    rowIndex,
                    column,
                    row.ColumnCount,
                    cell.Type,
                    cell.GetString(),
                    hasDouble,
                    hasDouble ? BitConverter.DoubleToInt64Bits(value) : 0));
            }
        }


        private static byte[] BuildLargeXlsx()
        {
            StringBuilder sb = new(512 * 1024);
            for (int r = 1; r <= 6000; r++)
            {
                sb.Append("<row r=\"").Append(r).Append("\">")
                  .Append("<c r=\"A").Append(r).Append("\"><v>").Append(r).Append("</v></c>")
                  .Append("<c r=\"B").Append(r).Append("\" t=\"inlineStr\"><is><t>row ").Append(r).Append(" text</t></is></c>")
                  .Append("<c r=\"C").Append(r).Append("\"><v>").Append(r * 1.5).Append("</v></c>")
                  .Append("<c r=\"D").Append(r).Append("\" t=\"b\"><v>").Append(r % 2).Append("</v></c>")
                  .Append("</row>");
            }
            using MemoryStream ms = WorkbookBuilder.Build(sb.ToString());
            return ms.ToArray();
        }

        private static byte[] BuildLargeXlsb()
        {
            return BuildLargeXlsbAsync().GetAwaiter().GetResult();
        }

        private static async Task<byte[]> BuildLargeXlsbAsync()
        {
            MemoryStream ms = new();
            await using (XlsbWorkbookWriter wb = XlsbWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsbSheetWriter sheet = wb.AddSheet("S1");
                await sheet.StartAsync();
                for (int r = 0; r < 6000; r++)
                {
                    await using XlsbRowWriter row = await sheet.StartRowAsync();
                    row.Write(r);
                    row.Write($"row {r} text");
                    row.Write(r * 1.5);
                    row.Write(r % 2 == 0);
                }
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.ToArray();
        }


        private static XlsxReader OpenXlsx(Stream stream, ExcelReaderOptions options)
        {
            return Excel.FromXlsx(stream, options: options);
        }

        private static XlsbReader OpenXlsb(Stream stream, ExcelReaderOptions options)
        {
            return Excel.FromXlsb(stream, options: options);
        }

        private static async ValueTask<IExcelRowReader> OpenXlsxAsync(Stream stream, ExcelReaderOptions options, CancellationToken ct)
        {
            return await Excel.FromXlsxAsync(stream, options: options, ct: ct);
        }

        private static async ValueTask<IExcelRowReader> OpenXlsbAsync(Stream stream, ExcelReaderOptions options, CancellationToken ct)
        {
            return await Excel.FromXlsbAsync(stream, options: options, ct: ct);
        }

        private static void CorruptEntryCompressedData(byte[] zipBytes, string entryName)
        {
            byte[] nameBytes = Encoding.UTF8.GetBytes(entryName);
            for (int i = 0; i + 46 <= zipBytes.Length; i++)
            {
                if (!IsCentralDirectoryEntryFor(zipBytes, i, nameBytes))
                {
                    continue;
                }
                ScrambleCompressedData(zipBytes, i);
                return;
            }
            throw new InvalidOperationException($"Central directory entry '{entryName}' not found.");
        }

        private static bool IsCentralDirectoryEntryFor(byte[] zipBytes, int i, byte[] nameBytes)
        {
            if (zipBytes[i] != 0x50 || zipBytes[i + 1] != 0x4B || zipBytes[i + 2] != 0x01 || zipBytes[i + 3] != 0x02)
            {
                return false;
            }
            int nameLen = BitConverter.ToUInt16(zipBytes, i + 28);
            if (nameLen != nameBytes.Length || i + 46 + nameLen > zipBytes.Length)
            {
                return false;
            }
            return zipBytes.AsSpan(i + 46, nameLen).SequenceEqual(nameBytes);
        }

        private static void ScrambleCompressedData(byte[] zipBytes, int centralDirectoryOffset)
        {
            uint compressedSize = BitConverter.ToUInt32(zipBytes, centralDirectoryOffset + 20);
            uint localHeaderOffset = BitConverter.ToUInt32(zipBytes, centralDirectoryOffset + 42);
            int localNameLen = BitConverter.ToUInt16(zipBytes, (int)localHeaderOffset + 26);
            int localExtraLen = BitConverter.ToUInt16(zipBytes, (int)localHeaderOffset + 28);
            int dataStart = (int)localHeaderOffset + 30 + localNameLen + localExtraLen;
            for (int k = 0; k < compressedSize; k++)
            {
                zipBytes[dataStart + k] = unchecked((byte)(0x5A ^ (k * 37)));
            }
        }

        public sealed record PrefetchFixture(
            string Name,
            Func<byte[]> Build,
            Func<Stream, ExcelReaderOptions, IExcelRowReader> OpenSync,
            Func<Stream, ExcelReaderOptions, CancellationToken, ValueTask<IExcelRowReader>> OpenAsync)
        {
            public override string ToString()
            {
                return Name;
            }
        }

    }
}
