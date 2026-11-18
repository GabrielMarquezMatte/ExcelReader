using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Internal;
using ExcelReader.Core.Writer.Xlsx;
using ExcelReader.Core.Writer.Zip;
using ExcelReader.Tests.Writer.Parallel;
using Item = ExcelReader.Tests.Writer.Parallel.ParallelRecordWriteTests.Item;

namespace ExcelReader.Tests.Writer.Zip
{
    public class DeflateTableReuseTests
    {
        public static TheoryData<int> Sizes()
        {
            return [0, 1, 200, 3_000, 40_000];
        }

        private static byte[] Segment(byte[] input)
        {
            using var output = new BiffBuffer(4096);
            DeflateEncoder.CompressSegment(input, output);
            return output.Span.ToArray();
        }

        private static byte[] Streamed(byte[] input)
        {
            return DeflateEncoderTests.Encode(input, static _ => 10_000);
        }

        private static T Cleared<T>(Func<T> encode)
        {
            DeflateEncoder.AlwaysClearTable = true;
            try
            {
                return encode();
            }
            finally
            {
                DeflateEncoder.AlwaysClearTable = false;
            }
        }

        [Theory]
        [MemberData(nameof(Sizes))]
        public void AReusedTableEncodesLikeAClearedOne(int rows)
        {
            byte[] input = DeflateEncoderTests.SheetXml(rows);
            byte[][] pieces = [.. input.Chunk(5_000), .. input.Chunk(70_000)];
            byte[][] expected = [.. pieces.Select(piece => Cleared(() => Segment(piece)))];
            byte[] streamed = Cleared(() => Streamed(input));
            for (int pass = 0; pass < 2; pass++)
            {
                for (int i = 0; i < pieces.Length; i++)
                {
                    Assert.Equal(expected[i], Segment(pieces[i]));
                }
                Assert.Equal(streamed, Streamed(input));
            }
        }

        [Theory]
        [InlineData((1u << 30) - 1)]
        [InlineData(uint.MaxValue - 5)]
        [InlineData(uint.MaxValue)]
        public void ABaseNearTheLimitEncodesLikeAClearedTable(uint tableBase)
        {
            byte[] a = DeflateEncoderTests.RandomBytes(50, seed: 1);
            byte[] c = DeflateEncoderTests.RandomBytes(100, seed: 3);
            // The match on the second copy of a skips the gram at 147, so only a stale slot can find it again at 400.
            byte[] input = [.. a, .. DeflateEncoderTests.RandomBytes(50, seed: 2), .. a, .. c,
                .. DeflateEncoderTests.RandomBytes(100, seed: 4), .. a[47..], c[0], .. DeflateEncoderTests.RandomBytes(50, seed: 5)];
            int shift = (int)Math.Min(0u - tableBase, 147u);
            byte[] earlier = [.. DeflateEncoderTests.RandomBytes(147 - shift, seed: 6), .. input[147..151], .. DeflateEncoderTests.RandomBytes(50, seed: 7)];
            byte[] segment = Cleared(() => Segment(input));
            byte[] streamed = Cleared(() => Streamed(input));

            Cleared(() => Segment(earlier));
            DeflateEncoder.SetThreadTableBase(tableBase);
            Assert.Equal(segment, Segment(input));
            Cleared(() => Segment(earlier));
            DeflateEncoder.SetThreadTableBase(tableBase);
            Assert.Equal(streamed, Streamed(input));
        }

        private static async Task<byte[]> WorkbookAsync(List<Item> items, bool parallel)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            var ms = new MemoryStream();
            // Shared-string numbering under parallel batches differs between runs even with a cleared table.
            await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true, new XlsxWriterOptions { UseSharedStrings = !parallel }))
            {
                XlsxSheetWriter sheet = wb.AddSheet("Data");
                if (parallel)
                {
                    await sheet.WriteRecordsParallelAsync(items, ParallelRecordWriteTests.Fill, new ParallelWriteOptions { DegreeOfParallelism = 4, BatchRows = 500 }, ct);
                }
                if (!parallel)
                {
                    await sheet.WriteRecordsAsync(items, (Action<XlsxRowWriter, Item>)ParallelRecordWriteTests.Fill, ct);
                }
                await sheet.EndAsync(ct);
                XlsxSheetWriter tiny = wb.AddSheet("Tiny");
                await tiny.WriteRecordsAsync(items.Take(3), (Action<XlsxRowWriter, Item>)ParallelRecordWriteTests.Fill, ct);
                await tiny.EndAsync(ct);
                await wb.EndAsync(ct);
            }
            return ZipLayout.WithoutTimestampsAndHost(ms.ToArray());
        }

        [Theory]
        [InlineData(10, false)]
        [InlineData(30_000, false)]
        [InlineData(30_000, true)]
        public async Task WorkbooksAreByteIdenticalToAClearedTable(int count, bool parallel)
        {
            List<Item> items = ParallelRecordWriteTests.Items(count);
            DeflateEncoder.AlwaysClearTable = true;
            byte[] expected;
            try
            {
                expected = await WorkbookAsync(items, parallel);
            }
            finally
            {
                DeflateEncoder.AlwaysClearTable = false;
            }

            Assert.Equal(expected, await WorkbookAsync(items, parallel));
            Assert.Equal(expected, await WorkbookAsync(items, parallel));
        }
    }
}
