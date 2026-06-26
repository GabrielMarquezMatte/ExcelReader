using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Reader
{
    // The Workbook OLE stream, read on demand instead of materialized. Two modes:
    //  - streamed: a seekable source + the stream's physical FAT-sector chain. Only one sector
    //    (plus a record-assembly scratch) is held at a time, so a 3 MB workbook costs ~KBs.
    //  - in-memory: a contiguous buffer (mini-stream workbooks, or a non-seekable fallback).
    // Immutable and shareable; each consumer reads through its own BiffCursor.
    [ExcludeFromCodeCoverage(Justification = "Exercised through XlsReader integration tests; guard-rail branches are corrupt-OLE only.")]
    internal sealed class WorkbookStream : IDisposable
    {
        private const int HeaderSize = 512;

        [SuppressMessage("SharpSource", "SS066:DisposableFieldIsNotDisposed",
            Justification = "Disposed in Dispose() when _ownsSource; otherwise the caller owns it.")]
        private readonly Stream? _source;
        private readonly bool _ownsSource;
        private readonly int[] _chain;        // physical sector numbers, in order
        private readonly ReadOnlyMemory<byte> _memory;

        internal int SectorSize { get; }
        internal long Length { get; }

        private WorkbookStream(Stream? source, bool ownsSource, int[] chain, ReadOnlyMemory<byte> memory, int sectorSize, long length)
        {
            _source = source;
            _ownsSource = ownsSource;
            _chain = chain;
            _memory = memory;
            SectorSize = sectorSize;
            Length = length;
        }

        internal static WorkbookStream Streamed(Stream source, bool ownsSource, int[] chain, int sectorSize, long length)
        {
            return new WorkbookStream(source, ownsSource, chain, default, sectorSize, length);
        }

        internal static WorkbookStream InMemory(ReadOnlyMemory<byte> data)
        {
            return new WorkbookStream(null, ownsSource: false, [], data, sectorSize: 1, data.Length);
        }

        internal BiffCursor OpenCursor()
        {
            return new BiffCursor(this);
        }

        internal bool IsMemory => _source is null;

        internal ReadOnlySpan<byte> Memory(long pos, int len)
        {
            return _memory.Span.Slice((int)pos, len);
        }

        // Reads sectorSize bytes of physical sector at chainIndex into dest (length == SectorSize).
        internal void LoadSector(int chainIndex, Span<byte> dest)
        {
            if ((uint)chainIndex >= (uint)_chain.Length)
            {
                throw new InvalidDataException("Invalid OLE sector chain index.");
            }
            long offset = HeaderSize + ((long)_chain[chainIndex] * SectorSize);
            _source!.Seek(offset, SeekOrigin.Begin);
            _source.ReadExactly(dest);
        }

        [SuppressMessage("IDisposableAnalyzers.Correctness", "IDISP007:Don't dispose injected",
            Justification = "Only disposed when _ownsSource — this type took ownership of the source in that case.")]
        public void Dispose()
        {
            if (_ownsSource)
            {
                _source?.Dispose();
            }
        }
    }
}
