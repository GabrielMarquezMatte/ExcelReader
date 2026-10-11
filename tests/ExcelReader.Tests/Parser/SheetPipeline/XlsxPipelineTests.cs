using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser.ParallelCsv;
using ExcelReader.Core.Parser.SheetPipeline;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Tests.Parser.SheetPipeline
{
    public class XlsxPipelineTests
    {
        private static readonly RowAggregation<List<string>> Collecting = new()
        {
            Seed = static () => [],
            Accumulate = static (ref List<string> rows, Row row) => rows.Add(Render(row)),
            Combine = static (left, right) =>
            {
                left.AddRange(right);
                return left;
            },
        };

        private static string Render(Row row)
        {
            var sb = new StringBuilder();
            foreach (RowCell cell in row.Cells)
            {
                sb.Append(cell.ColumnIndex).Append(':').Append(cell.Value.GetString()).Append('|');
            }
            return sb.ToString();
        }

        private static string Rows(int count, Func<int, string>? text = null)
        {
            var sb = new StringBuilder();
            for (int r = 1; r <= count; r++)
            {
                string value = text?.Invoke(r) ?? string.Create(CultureInfo.InvariantCulture, $"n{r}");
                sb.Append(CultureInfo.InvariantCulture, $"<row r=\"{r}\"><c r=\"A{r}\" t=\"inlineStr\"><is><t>{value}</t></is></c><c r=\"B{r}\"><v>{r}</v></c></row>");
            }
            return sb.ToString();
        }

        private static List<string> Sequential(XlsxWorkbook workbook, int headerRow)
        {
            var rows = new List<string>();
            foreach (Row row in workbook.FirstSheet)
            {
                rows.Add(Render(row));
            }
            return rows.Count > headerRow ? rows[headerRow..] : [];
        }

        private static Task<List<string>> PipelineAsync(XlsxWorkbook workbook, int headerRow, int dop, int chunkBytes, RowActionFactory<List<string>>? factory = null)
        {
            SheetPipelineSource source = SheetPipelineSource.Create(workbook.FirstSheet)!;
            return SheetPipelineRunner.AggregateAsync(source, Collecting, factory, headerRow, dop, chunkBytes, TestContext.Current.CancellationToken);
        }

        public static TheoryData<int, int> Shapes()
        {
            var data = new TheoryData<int, int>();
            foreach (int dop in new[] { 2, 4, 8 })
            {
                foreach (int chunk in new[] { 64, 200, 1000, 1 << 20 })
                {
                    data.Add(dop, chunk);
                }
            }
            return data;
        }

        [Theory]
        [MemberData(nameof(Shapes))]
        public async Task MatchesSequential(int dop, int chunkBytes)
        {
            using MemoryStream ms = WorkbookBuilder.Build(Rows(500));
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);

            Assert.Equal(Sequential(workbook, 0), await PipelineAsync(workbook, 0, dop, chunkBytes));
            Assert.Equal(Sequential(workbook, 1), await PipelineAsync(workbook, 1, dop, chunkBytes));
            Assert.Equal(Sequential(workbook, 3), await PipelineAsync(workbook, 3, dop, chunkBytes));
        }

        [Fact]
        public async Task SourceStreamCancellationIsNotSwallowed()
        {
            using var ms = WorkbookBuilder.Build(Rows(20000));
            using var stream = new CancelAfterArmedStream(ms);
            using XlsxWorkbook workbook = Excel.FromXlsx(stream);
            SheetPipelineSource source = SheetPipelineSource.Create(workbook.FirstSheet)!;
            stream.Arm(100 * 1024);

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                SheetPipelineRunner.AggregateAsync(source, Collecting, null, 0, 2, 64, CancellationToken.None));
        }

        private sealed class CancelAfterArmedStream(Stream inner) : Stream
        {
            private long _remaining = long.MaxValue;

            public void Arm(long budget) => _remaining = budget;

            public override bool CanRead => inner.CanRead;
            public override bool CanSeek => inner.CanSeek;
            public override bool CanWrite => false;
            public override long Length => inner.Length;
            public override long Position
            {
                get => inner.Position;
                set => inner.Position = value;
            }

            public override void Flush()
            {
            }

            public override int Read(byte[] buffer, int offset, int count)
            {
                int n = inner.Read(buffer, offset, count);
                Spend(n);
                return n;
            }

            public override int Read(Span<byte> buffer)
            {
                int n = inner.Read(buffer);
                Spend(n);
                return n;
            }

            public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
            {
                int n = await inner.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
                Spend(n);
                return n;
            }

            public override long Seek(long offset, SeekOrigin origin) => inner.Seek(offset, origin);
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

            private void Spend(int n)
            {
                _remaining -= n;
                if (_remaining < 0)
                {
                    throw new OperationCanceledException("source timeout");
                }
            }
        }

        [Theory]
        [InlineData("")]
        [InlineData("<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>")]
        public async Task ShortSheetsYieldNothing(string rows)
        {
            using MemoryStream ms = WorkbookBuilder.Build(rows);
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);

            Assert.Empty(await PipelineAsync(workbook, 1, 4, 64));
            Assert.Empty(await PipelineAsync(workbook, 5, 4, 64));
        }

        [Fact]
        public async Task TheHeaderRowReachesTheFactoryOnce()
        {
            using MemoryStream ms = WorkbookBuilder.Build(Rows(300));
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            int calls = 0;
            string? header = null;
            RowActionFactory<List<string>> factory = (Row row, bool sequential) =>
            {
                Interlocked.Increment(ref calls);
                header = Render(row);
                Assert.False(sequential);
                return Collecting.Accumulate;
            };

            List<string> rows = await PipelineAsync(workbook, 2, 4, 128, factory);

            Assert.Equal(1, calls);
            Assert.Equal(Sequential(workbook, 0)[1], header);
            Assert.Equal(Sequential(workbook, 2), rows);
        }

        [Fact]
        public async Task APrefixedSheetMatchesSequential()
        {
            var sb = new StringBuilder();
            for (int r = 1; r <= 200; r++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"<x:row r=\"{r}\"><x:c r=\"A{r}\"><x:v>{r}</x:v></x:c></x:row>");
            }
            using MemoryStream ms = WorkbookBuilder.BuildPrefixed("x", sb.ToString());
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);

            Assert.Equal(Sequential(workbook, 0), await PipelineAsync(workbook, 0, 4, 96));
        }

        [Theory]
        [InlineData("<![CDATA[ <row r=\"9\"> ]]>")]
        [InlineData("</t></is></c><!-- <row r=\"9\"> --><c t=\"inlineStr\"><is><t>z")]
        public async Task ARowLikeTagInsideARowMatchesSequential(string poison)
        {
            using MemoryStream ms = WorkbookBuilder.Build(Rows(120, r => r % 10 == 0 ? poison : "n"));
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            List<string> expected = Sequential(workbook, 0);

            foreach (int chunkBytes in new[] { 64, 90, 150, 400 })
            {
                Assert.Equal(expected, await PipelineAsync(workbook, 0, 4, chunkBytes));
                Assert.Equal(expected[1..], await PipelineAsync(workbook, 1, 4, chunkBytes));
            }
        }

        [Fact]
        public async Task RowsAfterTheEndOfSheetDataAreIgnored()
        {
            string rows = Rows(60) + "</sheetData><extLst>" + Rows(60) + "</extLst><sheetData>";
            using MemoryStream ms = WorkbookBuilder.Build(rows);
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            List<string> expected = Sequential(workbook, 0);

            Assert.Equal(60, expected.Count);
            Assert.Equal(expected, await PipelineAsync(workbook, 0, 4, 80));
        }

        [Fact]
        public async Task AnAccumulateExceptionSurfaces()
        {
            using MemoryStream ms = WorkbookBuilder.Build(Rows(300));
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            var failing = new RowAggregation<int>
            {
                Seed = static () => 0,
                Accumulate = static (ref int count, Row row) =>
                {
                    if (row[1].TryGetDouble(out double value) && value == 150)
                    {
                        throw new InvalidOperationException("boom");
                    }
                    count++;
                },
                Combine = static (left, right) => left + right,
            };
            SheetPipelineSource source = SheetPipelineSource.Create(workbook.FirstSheet)!;

            InvalidOperationException ex = await Assert.ThrowsAsync<InvalidOperationException>(
                () => SheetPipelineRunner.AggregateAsync(source, failing, null, 0, 4, 128, TestContext.Current.CancellationToken));

            Assert.Equal("boom", ex.Message);
        }

        [Fact]
        public async Task CancellationStopsTheRead()
        {
            using MemoryStream ms = WorkbookBuilder.Build(Rows(2000));
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(TestContext.Current.CancellationToken);
            int seen = 0;
            var cancelling = new RowAggregation<int>
            {
                Seed = static () => 0,
                Accumulate = (ref int count, Row _) =>
                {
                    count++;
                    if (Interlocked.Increment(ref seen) == 5)
                    {
                        cts.Cancel();
                    }
                },
                Combine = static (left, right) => left + right,
            };
            SheetPipelineSource source = SheetPipelineSource.Create(workbook.FirstSheet)!;

            await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => SheetPipelineRunner.AggregateAsync(source, cancelling, null, 0, 4, 64, cts.Token));
        }
    }
}
