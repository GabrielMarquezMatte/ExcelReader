using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Tests.Reader.Xlsx
{
    public class XlsxAsyncRowBoundaryTests
    {
        private const int RowCount = 4_000;

        private sealed class YieldingStream(byte[] bytes) : Stream
        {
            private const int MaxPerRead = 4096;

            private long _position;

            public override bool CanRead
            {
                get
                {
                    return true;
                }
            }

            public override bool CanSeek
            {
                get
                {
                    return true;
                }
            }

            public override bool CanWrite
            {
                get
                {
                    return false;
                }
            }

            public override long Length
            {
                get
                {
                    return bytes.Length;
                }
            }

            public override long Position
            {
                get
                {
                    return _position;
                }
                set
                {
                    _position = value;
                }
            }

            public override int Read(byte[] buffer, int offset, int count)
            {
                return Read(buffer.AsSpan(offset, count));
            }

            public override int Read(Span<byte> buffer)
            {
                int available = (int)Math.Max(0, Math.Min(bytes.Length - _position, Math.Min(buffer.Length, MaxPerRead)));
                bytes.AsSpan((int)_position, available).CopyTo(buffer);
                _position += available;
                return available;
            }

            public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
            {
                await Task.Yield();
                return Read(buffer.Span);
            }

            public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
            {
                return ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
            }

            public override long Seek(long offset, SeekOrigin origin)
            {
                _position = origin switch
                {
                    SeekOrigin.Begin => offset,
                    SeekOrigin.Current => _position + offset,
                    _ => bytes.Length + offset,
                };
                return _position;
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
        }

        [Fact]
        public async Task AsyncReadKeepsRowsWhoseOpenTagStraddlesARefill()
        {
            byte[] bytes = BuildWideRowTagWorkbook();
            string[] expected = [.. Enumerable.Range(1, RowCount).Select(static i => i.ToString(CultureInfo.InvariantCulture))];

            await using XlsxWorkbook workbook = await Excel.FromXlsxAsync(
                new YieldingStream(bytes), leaveOpen: false, ct: TestContext.Current.CancellationToken);
            List<string> values = [];
            await using (XlsxWorkbook.Enumerator rows = workbook.FirstSheet.GetAsyncEnumerator(TestContext.Current.CancellationToken))
            {
                while (await rows.MoveNextAsync())
                {
                    values.Add(rows.Current[0].GetString());
                }
            }

            Assert.Equal(expected, values, StringComparer.Ordinal);
            Assert.Equal(expected, ReadSync(bytes), StringComparer.Ordinal);
        }

        private static List<string> ReadSync(byte[] bytes)
        {
            using XlsxWorkbook workbook = Excel.FromXlsx(new YieldingStream(bytes), leaveOpen: false);
            using XlsxWorkbook.Enumerator rows = workbook.FirstSheet.GetEnumerator();
            List<string> values = [];
            while (rows.MoveNext())
            {
                values.Add(rows.Current[0].GetString());
            }
            return values;
        }

        private static byte[] BuildWideRowTagWorkbook()
        {
            string padding = new('p', 400);
            StringBuilder rows = new();
            for (int i = 1; i <= RowCount; i++)
            {
                rows.Append(CultureInfo.InvariantCulture,
                    $"""<row r="{i}" spans="1:1" pad="{padding}"><c r="A{i}"><v>{i}</v></c></row>""");
            }
            using MemoryStream compressed = WorkbookBuilder.Build(rows.ToString());
            return Stored(compressed);
        }

        private static byte[] Stored(MemoryStream compressed)
        {
            using MemoryStream stored = new();
            using (ZipArchive source = new(compressed, ZipArchiveMode.Read))
            using (ZipArchive target = new(stored, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (ZipArchiveEntry entry in source.Entries)
                {
                    using Stream from = entry.Open();
                    using Stream to = target.CreateEntry(entry.FullName, CompressionLevel.NoCompression).Open();
                    from.CopyTo(to);
                }
            }
            return stored.ToArray();
        }
    }
}
