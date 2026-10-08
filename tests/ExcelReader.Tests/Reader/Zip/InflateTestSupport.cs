using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Reader.Zip.Inflate;

namespace ExcelReader.Tests.Reader.Zip
{
    [SuppressMessage("Security", "CA5394:Do not use insecure randomness", Justification = "Seeded test data, not security.")]
    internal static class InflateTestSupport
    {
        internal static byte[] Deflate(byte[] raw, CompressionLevel level)
        {
            using var compressed = new MemoryStream();
            using (var deflate = new DeflateStream(compressed, level, leaveOpen: true))
            {
                deflate.Write(raw);
            }
            return compressed.ToArray();
        }

        internal static byte[] Inflate(byte[] compressed, int readSize = 4096, int sourceStep = int.MaxValue)
        {
            using var inflate = new InflateStream(new SteppedStream(compressed, sourceStep));
            using var output = new MemoryStream();
            byte[] buffer = new byte[readSize];
            int read;
            while ((read = inflate.Read(buffer)) > 0)
            {
                output.Write(buffer, 0, read);
            }
            return output.ToArray();
        }

        internal static async Task<byte[]> InflateAsync(byte[] compressed, int readSize, int sourceStep, CancellationToken ct)
        {
            using var inflate = new InflateStream(new SteppedStream(compressed, sourceStep));
            using var output = new MemoryStream();
            byte[] buffer = new byte[readSize];
            int read;
            while ((read = await inflate.ReadAsync(buffer, ct)) > 0)
            {
                output.Write(buffer, 0, read);
            }
            return output.ToArray();
        }

        internal static byte[] Generate(DataShape shape, int size, int seed)
        {
            var random = new Random(seed);
            byte[] data = new byte[size];
            switch (shape)
            {
                case DataShape.Zeros:
                    break;
                case DataShape.Random:
                    random.NextBytes(data);
                    break;
                case DataShape.SheetXml:
                    FillSheetXml(data, random);
                    break;
                case DataShape.ShortPeriods:
                    FillShortPeriods(data, random);
                    break;
                case DataShape.RepeatedBlock:
                    FillRepeatedBlock(data, random);
                    break;
                default:
                    FillMixedRuns(data, random);
                    break;
            }
            return data;
        }

        private static void FillSheetXml(byte[] data, Random random)
        {
            var xml = new StringBuilder();
            for (int row = 1; xml.Length < data.Length; row++)
            {
                xml.Append(System.Globalization.CultureInfo.InvariantCulture,
                    $"<row r=\"{row}\"><c r=\"A{row}\"><v>{random.NextDouble() * 1000:F4}</v></c><c r=\"B{row}\" t=\"s\"><v>{random.Next(500)}</v></c></row>");
            }
            Encoding.ASCII.GetBytes(xml.ToString(0, data.Length), data);
        }

        // Runs whose period is 1 to 9 bytes, so matches land on every short offset.
        private static void FillShortPeriods(byte[] data, Random random)
        {
            int position = 0;
            while (position < data.Length)
            {
                int period = random.Next(1, 10);
                int run = Math.Min(data.Length - position, random.Next(period, 600));
                for (int i = 0; i < run; i++)
                {
                    data[position + i] = i < period ? (byte)random.Next(256) : data[position + i - period];
                }
                position += run;
            }
        }

        // One incompressible 32,768-byte block repeated, so matches sit at the far edge of the window.
        private static void FillRepeatedBlock(byte[] data, Random random)
        {
            const int Block = 32768;
            random.NextBytes(data.AsSpan(0, Math.Min(Block, data.Length)));
            for (int i = Block; i < data.Length; i++)
            {
                data[i] = data[i - Block];
            }
        }

        private static void FillMixedRuns(byte[] data, Random random)
        {
            int position = 0;
            while (position < data.Length)
            {
                int run = Math.Min(data.Length - position, random.Next(1, 400));
                if (random.Next(3) == 0)
                {
                    random.NextBytes(data.AsSpan(position, run));
                }
                else
                {
                    data.AsSpan(position, run).Fill((byte)random.Next(256));
                }
                position += run;
            }
        }
    }

    public enum DataShape
    {
        Zeros,
        Random,
        SheetXml,
        ShortPeriods,
        RepeatedBlock,
        MixedRuns,
    }

    internal sealed class SteppedStream(byte[] data, int step) : Stream
    {
        private int _position;

        internal bool Disposed { get; private set; }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => data.Length;

        public override long Position
        {
            get => _position;
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return Read(buffer.AsSpan(offset, count));
        }

        public override int Read(Span<byte> buffer)
        {
            int count = Math.Min(Math.Min(step, buffer.Length), data.Length - _position);
            data.AsSpan(_position, count).CopyTo(buffer);
            _position += count;
            return count;
        }

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            await Task.Yield();
            return Read(buffer.Span);
        }

        public override void Flush()
        {
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
            throw new NotSupportedException();
        }

        protected override void Dispose(bool disposing)
        {
            Disposed = true;
            base.Dispose(disposing);
        }
    }
}
