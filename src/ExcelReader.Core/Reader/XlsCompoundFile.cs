using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using ExcelReader.Core.Crypto;

namespace ExcelReader.Core.Reader
{
    // Parses the OLE/CFB container metadata (header, FAT, directory, mini-FAT/stream) via
    // CfbContainer, then locates the Workbook stream and hands it back as a WorkbookStream that
    // reads its sectors on demand. Non-seekable sources are buffered into a MemoryStream first
    // (rare fallback, same cost as before).
    [ExcludeFromCodeCoverage(Justification = "Covered through XlsReader integration tests; most uncovered paths are corrupt-OLE guard rails.")]
    internal static class XlsCompoundFile
    {
        internal static ReadOnlySpan<byte> Signature => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

        internal static WorkbookStream OpenWorkbook(Stream stream, bool leaveOpen, ExcelReaderOptions? options = null)
        {
            (Stream source, bool ownsSource) = EnsureSeekable(stream, leaveOpen);
            try
            {
                return BuildWorkbook(source, ownsSource, options ?? ExcelReaderOptions.Default);
            }
            catch
            {
                if (ownsSource)
                {
                    source.Dispose();
                }
                throw;
            }
        }

        internal static WorkbookStream OpenWorkbook(ReadOnlyMemory<byte> data, ExcelReaderOptions? options = null)
        {
            using MemoryStream metadata = AsStream(data);
            return BuildWorkbook(metadata, ownsSource: false, options ?? ExcelReaderOptions.Default, memory: data);
        }

        internal static async ValueTask<WorkbookStream> OpenWorkbookAsync(Stream stream, bool leaveOpen, ExcelReaderOptions? options, CancellationToken ct)
        {
            (Stream source, bool ownsSource) = await EnsureSeekableAsync(stream, leaveOpen, ct).ConfigureAwait(false);
            try
            {
                return BuildWorkbook(source, ownsSource, options ?? ExcelReaderOptions.Default);
            }
            catch
            {
                if (ownsSource)
                {
                    await source.DisposeAsync().ConfigureAwait(false);
                }
                throw;
            }
        }

        private static (Stream Source, bool OwnsSource) EnsureSeekable(Stream stream, bool leaveOpen)
        {
            if (stream.CanSeek)
            {
                return (stream, !leaveOpen);
            }
            MemoryStream ms = new();
            stream.CopyTo(ms);
            if (!leaveOpen)
            {
                stream.Dispose();
            }
            ms.Position = 0;
            return (ms, true);
        }

        private static async ValueTask<(Stream Source, bool OwnsSource)> EnsureSeekableAsync(Stream stream, bool leaveOpen, CancellationToken ct)
        {
            if (stream.CanSeek)
            {
                return (stream, !leaveOpen);
            }
            MemoryStream ms = new();
            await stream.CopyToAsync(ms, ct).ConfigureAwait(false);
            if (!leaveOpen)
            {
                await stream.DisposeAsync().ConfigureAwait(false);
            }
            ms.Position = 0;
            return (ms, true);
        }

        internal static MemoryStream AsStream(ReadOnlyMemory<byte> data)
        {
            if (MemoryMarshal.TryGetArray(data, out ArraySegment<byte> segment))
            {
                return new MemoryStream(segment.Array!, segment.Offset, segment.Count, writable: false);
            }
            return new MemoryStream(data.ToArray(), writable: false);
        }

        private static WorkbookStream BuildWorkbook(Stream source, bool ownsSource, ExcelReaderOptions options, ReadOnlyMemory<byte> memory = default)
        {
            CfbContainer cfb = CfbContainer.Parse(source, ownsSource, options, memory);
            try
            {
                // "Workbook" is BIFF8; "Book" is BIFF5 - both mean the same stream to this reader.
                string name;
                if (cfb.ContainsStream("Workbook"))
                {
                    name = "Workbook";
                }
                else if (cfb.ContainsStream("Book"))
                {
                    name = "Book";
                }
                else
                {
                    throw new InvalidDataException("The OLE document does not contain a Workbook stream.");
                }
                cfb.TryFindEntry(name, out CfbContainer.DirectoryEntry workbook);

                // A stream cannot hold more content than the container's own byte length, so an
                // inflated Size field (the same attack class as fatSectorCount/difatSectorCount above)
                // is a crafted header — reject it before it drives an allocation or a chain walk sized
                // off it. The caller's byte budget applies here too, since this is the one choke point
                // both the mini-stream and chained/streamed branches below pass through.
                if (workbook.Size < 0 || workbook.Size > source.Length)
                {
                    throw new InvalidDataException("The OLE Workbook stream size exceeds the container.");
                }
                LimitChecks.ThrowIfEntryLengthExceeds(workbook.Size, options.MaxTotalDecompressedBytes, nameof(ExcelReaderOptions.MaxTotalDecompressedBytes));

                // Mini-stream workbooks (tiny, rare) are materialized; everything else streams.
                if (workbook.Size < cfb.MiniCutoff && workbook.StartSector >= 0)
                {
                    byte[] data = cfb.ReadStream(name, options.MaxTotalDecompressedBytes);
                    if (ownsSource)
                    {
                        source.Dispose();
                    }
                    return WorkbookStream.InMemory(data);
                }

                int chainCount = CfbContainer.SectorCount(workbook.Size, cfb.SectorSize);
                int[] chain = CfbContainer.BuildChain(cfb.Fat, workbook.StartSector, chainCount);
                if (!memory.IsEmpty)
                {
                    return WorkbookStream.Chained(memory, chain, chainCount, cfb.SectorSize, workbook.Size);
                }
                return WorkbookStream.Streamed(source, ownsSource, chain, chainCount, cfb.SectorSize, workbook.Size);
            }
            finally
            {
                // Only return the pooled FAT array here. Disposal of `source` past this point is still
                // governed by the `ownsSource` flag each branch above already handles explicitly
                // (mini-stream disposes it itself; Chained/Streamed hand ownership off to WorkbookStream),
                // so calling the container's full Dispose() here would double-dispose it.
                cfb.ReturnFatBuffer();
            }
        }
    }
}
