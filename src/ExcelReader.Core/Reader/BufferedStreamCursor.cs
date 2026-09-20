using System.Buffers;
using System.Runtime.InteropServices;

namespace ExcelReader.Core.Reader
{
    internal sealed class BufferedStreamCursor
    {
        private const int InitialBuf = 64 * 1024;

        private readonly int _maxCellBytes;
        private readonly string _limitName;

        private readonly bool _pooled;

        internal byte[] Buf { get; private set; }
        internal int Pos { get; set; }
        internal int Len { get; private set; }
        internal bool Eof { get; private set; }

        internal long BaseOffset { get; private set; }

        internal BufferedStreamCursor(int maxCellBytes, string limitName, int initialCapacity = InitialBuf)
        {
            _maxCellBytes = maxCellBytes;
            _limitName = limitName;
            _pooled = true;
            Buf = ArrayPool<byte>.Shared.Rent(initialCapacity);
        }

        internal BufferedStreamCursor(ReadOnlyMemory<byte> content, int maxCellBytes, string limitName)
        {
            _maxCellBytes = maxCellBytes;
            _limitName = limitName;
            _pooled = false;
            Eof = true;
            if (MemoryMarshal.TryGetArray(content, out ArraySegment<byte> segment))
            {
                Buf = segment.Array!;
                Pos = segment.Offset;
                Len = segment.Offset + segment.Count;
                BaseOffset = -segment.Offset;
                return;
            }
            Buf = content.ToArray();
            Len = Buf.Length;
        }

        private void PrepareBuffer()
        {
            int freeTail = Buf.Length - Len;
            if (freeTail > 0 && (Pos == 0 || freeTail >= Buf.Length / 4))
            {
                return;
            }
            if (Pos > 0)
            {
                Buf.AsSpan(Pos, Len - Pos).CopyTo(Buf);
                Len -= Pos;
                BaseOffset += Pos;
                Pos = 0;
                return;
            }
            byte[] bigger = ArrayPool<byte>.Shared.Rent(LimitChecks.NextBufferSize(_maxCellBytes, _limitName, Buf.Length, Buf.Length + 1));
            Buf.AsSpan(0, Len).CopyTo(bigger);
            ArrayPool<byte>.Shared.Return(Buf);
            Buf = bigger;
        }

        internal void Fill(Stream? source)
        {
            if (source is null)
            {
                return;
            }
            PrepareBuffer();
            int n = source.Read(Buf, Len, Buf.Length - Len);
            if (n == 0)
            {
                Eof = true;
                return;
            }
            Len += n;
        }

        internal ValueTask FillAsync(Stream? source, CancellationToken ct)
        {
            return source is null ? ValueTask.CompletedTask : FillFromAsync(source, ct);
        }

        private async ValueTask FillFromAsync(Stream source, CancellationToken ct)
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

        internal void Ensure(Stream? source, int n)
        {
            if (Len - Pos >= n || Eof)
            {
                return;
            }
            EnsureSlow(source, n);
        }

        private void EnsureSlow(Stream? source, int n)
        {
            while (Len - Pos < n && !Eof)
            {
                Fill(source);
            }
        }

        internal ValueTask EnsureAsync(Stream? source, int n, CancellationToken ct)
        {
            if (Len - Pos >= n || Eof)
            {
                return ValueTask.CompletedTask;
            }
            return EnsureSlowAsync(source, n, ct);
        }

        private async ValueTask EnsureSlowAsync(Stream? source, int n, CancellationToken ct)
        {
            while (Len - Pos < n && !Eof)
            {
                await FillAsync(source, ct).ConfigureAwait(false);
            }
        }

        internal void Return()
        {
            if (_pooled && Buf.Length > 0)
            {
                ArrayPool<byte>.Shared.Return(Buf);
            }
            Buf = [];
        }
    }
}
