using System.Buffers;

namespace ExcelReader.Core.Reader
{
    // Pooled-buffer refill/compact-or-grow cursor shared by the XLSX/XLSB/CSV forward-only stream
    // enumerators. Growth is capped the same way CellAccumulator caps its value buffer
    // (LimitChecks.NextBufferSize); each caller supplies its own (maxCellBytes, limitName) pair since
    // XLSX/XLSB draw the limit from ExcelReaderOptions while CSV draws it from CsvReaderOptions.
    internal sealed class BufferedStreamCursor
    {
        private const int InitialBuf = 64 * 1024;

        private readonly int _maxCellBytes;
        private readonly string _limitName;

        internal byte[] Buf { get; private set; }
        internal int Pos { get; set; }
        internal int Len { get; private set; }
        internal bool Eof { get; private set; }

        internal BufferedStreamCursor(int maxCellBytes, string limitName)
        {
            _maxCellBytes = maxCellBytes;
            _limitName = limitName;
            Buf = ArrayPool<byte>.Shared.Rent(InitialBuf);
        }

        // Compact the consumed prefix, or grow the buffer if every byte is still unprocessed.
        private void PrepareBuffer()
        {
            if (Pos > 0)
            {
                Buf.AsSpan(Pos, Len - Pos).CopyTo(Buf);
                Len -= Pos;
                Pos = 0;
                return;
            }
            if (Len != Buf.Length)
            {
                return;
            }
            byte[] bigger = ArrayPool<byte>.Shared.Rent(LimitChecks.NextBufferSize(_maxCellBytes, _limitName, Buf.Length, Buf.Length + 1));
            Buf.AsSpan(0, Len).CopyTo(bigger);
            ArrayPool<byte>.Shared.Return(Buf);
            Buf = bigger;
        }

        internal void Fill(Stream source)
        {
            PrepareBuffer();
            int n = source.Read(Buf, Len, Buf.Length - Len);
            if (n == 0)
            {
                Eof = true;
                return;
            }
            Len += n;
        }

        internal async ValueTask FillAsync(Stream source, CancellationToken ct)
        {
            PrepareBuffer();
            int n = await source.ReadAsync(Buf.AsMemory(Len, Buf.Length - Len), ct).ConfigureAwait(false);
            if (n == 0)
            {
                Eof = true;
                return;
            }
            Len += n;
        }

        internal void Ensure(Stream source, int n)
        {
            while (Len - Pos < n && !Eof)
            {
                Fill(source);
            }
        }

        internal ValueTask EnsureAsync(Stream source, int n, CancellationToken ct)
        {
            if (Len - Pos >= n || Eof)
            {
                return ValueTask.CompletedTask;
            }
            return EnsureSlowAsync(source, n, ct);
        }

        private async ValueTask EnsureSlowAsync(Stream source, int n, CancellationToken ct)
        {
            while (Len - Pos < n && !Eof)
            {
                await FillAsync(source, ct).ConfigureAwait(false);
            }
        }

        internal void Return()
        {
            if (Buf.Length > 0)
            {
                ArrayPool<byte>.Shared.Return(Buf);
                Buf = [];
            }
        }
    }
}
