using System.Buffers.Binary;

namespace ExcelReader.Core.Writer.Internal
{
    /// <summary>One named stream in a CFB container: its size, and how to write its bytes.</summary>
    internal readonly record struct CfbStreamSpec(
        string Name,
        long Size,
        Action<Stream> WriteBody,
        Func<Stream, CancellationToken, ValueTask> WriteBodyAsync);

    internal static partial class OleCompoundWriter
    {
        private const int HeaderSize = 512;
        private const int SectorSize = 512;
        private const int MiniSectorSize = 64;
        private const int MiniCutoff = 4096;
        private const int DirectoryEntrySize = 128;
        private const int DirectoryEntriesPerSector = SectorSize / DirectoryEntrySize;
        private const int EndOfChain = unchecked((int)0xFFFFFFFE);
        private const int FatSectorMarker = unchecked((int)0xFFFFFFFD);
        private const int DifatSectorMarker = unchecked((int)0xFFFFFFFC);
        private const int FreeSector = unchecked((int)0xFFFFFFFF);
        private const int NoStream = unchecked((int)0xFFFFFFFF);
        private const int FatEntriesPerSector = SectorSize / 4;
        private const int MaxHeaderDifat = (HeaderSize - 0x4C) / 4;
        private const int DifatEntriesPerSector = FatEntriesPerSector - 1;

        internal static void Write(Stream destination, IReadOnlyList<CfbStreamSpec> streams)
        {
            Layout layout = Layout.Compute(streams);
            destination.Write(layout.Header);
            destination.Write(layout.Fat);
            if (layout.Difat.Length > 0)
            {
                destination.Write(layout.Difat);
            }
            destination.Write(layout.Directory);
            if (layout.MiniFat.Length > 0)
            {
                destination.Write(layout.MiniFat);
            }
            for (int k = 0; k < layout.MiniOrder.Length; k++)
            {
                streams[layout.MiniOrder[k]].WriteBody(destination);
                WritePadding(destination, layout.MiniPadding[k]);
            }
            WritePadding(destination, layout.MiniBlobPadding);
            for (int k = 0; k < layout.BigOrder.Length; k++)
            {
                streams[layout.BigOrder[k]].WriteBody(destination);
                WritePadding(destination, layout.BigPadding[k]);
            }
        }

        internal static async ValueTask WriteAsync(Stream destination, IReadOnlyList<CfbStreamSpec> streams, CancellationToken ct)
        {
            Layout layout = Layout.Compute(streams);
            await destination.WriteAsync(layout.Header, ct).ConfigureAwait(false);
            await destination.WriteAsync(layout.Fat, ct).ConfigureAwait(false);
            if (layout.Difat.Length > 0)
            {
                await destination.WriteAsync(layout.Difat, ct).ConfigureAwait(false);
            }
            await destination.WriteAsync(layout.Directory, ct).ConfigureAwait(false);
            if (layout.MiniFat.Length > 0)
            {
                await destination.WriteAsync(layout.MiniFat, ct).ConfigureAwait(false);
            }
            for (int k = 0; k < layout.MiniOrder.Length; k++)
            {
                await streams[layout.MiniOrder[k]].WriteBodyAsync(destination, ct).ConfigureAwait(false);
                await WritePaddingAsync(destination, layout.MiniPadding[k], ct).ConfigureAwait(false);
            }
            await WritePaddingAsync(destination, layout.MiniBlobPadding, ct).ConfigureAwait(false);
            for (int k = 0; k < layout.BigOrder.Length; k++)
            {
                await streams[layout.BigOrder[k]].WriteBodyAsync(destination, ct).ConfigureAwait(false);
                await WritePaddingAsync(destination, layout.BigPadding[k], ct).ConfigureAwait(false);
            }
        }

        private static void WritePadding(Stream destination, int count)
        {
            if (count > 0)
            {
                destination.Write(new byte[count]);
            }
        }

        private static async ValueTask WritePaddingAsync(Stream destination, int count, CancellationToken ct)
        {
            if (count > 0)
            {
                await destination.WriteAsync(new byte[count], ct).ConfigureAwait(false);
            }
        }

        private static int CeilingDiv(int n, int d)
        {
            return (n + d - 1) / d;
        }

        private static int CeilingDivLong(long n, int d)
        {
            long result = (n + d - 1) / d;
            if (result > int.MaxValue)
            {
                throw new ArgumentOutOfRangeException(nameof(n), "The CFB container would need more sectors than [MS-CFB] can address.");
            }
            return (int)result;
        }

        private static void WriteU16(Span<byte> dest, int offset, ushort value)
        {
            BinaryPrimitives.WriteUInt16LittleEndian(dest.Slice(offset, 2), value);
        }

        private static void WriteI32(Span<byte> dest, int offset, int value)
        {
            BinaryPrimitives.WriteInt32LittleEndian(dest.Slice(offset, 4), value);
        }

        private sealed class Layout
        {
            internal byte[] Header = [];
            internal byte[] Fat = [];
            internal byte[] Difat = [];
            internal byte[] Directory = [];
            internal byte[] MiniFat = [];
            internal int[] MiniOrder = [];
            internal int[] BigOrder = [];
            internal int[] MiniPadding = [];
            internal int[] BigPadding = [];
            internal int MiniBlobPadding;

            internal static Layout Compute(IReadOnlyList<CfbStreamSpec> streams)
            {
                ArgumentNullException.ThrowIfNull(streams);
                if (streams.Count == 0)
                {
                    throw new ArgumentException("A CFB container needs at least one stream.", nameof(streams));
                }
                foreach (CfbStreamSpec spec in streams)
                {
                    if (spec.Size <= 0)
                    {
                        throw new ArgumentException($"Stream '{spec.Name}' must have a positive size.", nameof(streams));
                    }
                }

                var layout = new Layout();
                List<int> mini = [];
                List<int> big = [];
                for (int i = 0; i < streams.Count; i++)
                {
                    (streams[i].Size < MiniCutoff ? mini : big).Add(i);
                }
                layout.MiniOrder = [.. mini];
                layout.BigOrder = [.. big];

                int[] firstMiniSector = new int[mini.Count];
                layout.MiniPadding = new int[mini.Count];
                int miniSectors = 0;
                for (int k = 0; k < mini.Count; k++)
                {
                    long size = streams[mini[k]].Size;
                    int sectors = CeilingDivLong(size, MiniSectorSize);
                    firstMiniSector[k] = miniSectors;
                    layout.MiniPadding[k] = (int)((long)sectors * MiniSectorSize - size);
                    miniSectors += sectors;
                }
                int miniBlobBytes = miniSectors * MiniSectorSize;
                int miniStreamSectors = CeilingDiv(miniBlobBytes, SectorSize);
                layout.MiniBlobPadding = (miniStreamSectors * SectorSize) - miniBlobBytes;
                int miniFatSectors = miniSectors == 0 ? 0 : CeilingDiv(miniSectors, FatEntriesPerSector);

                int[] bigSectorCount = new int[big.Count];
                layout.BigPadding = new int[big.Count];
                int bigSectorTotal = 0;
                for (int k = 0; k < big.Count; k++)
                {
                    long size = streams[big[k]].Size;
                    int sectors = CeilingDivLong(size, SectorSize);
                    bigSectorCount[k] = sectors;
                    layout.BigPadding[k] = (int)((long)sectors * SectorSize - size);
                    bigSectorTotal += sectors;
                }

                int directorySectors = CeilingDiv(streams.Count + 1, DirectoryEntriesPerSector);
                int payloadSectors = directorySectors + miniFatSectors + miniStreamSectors + bigSectorTotal;
                (int fatCount, int difatCount) = ComputeSectorCounts(payloadSectors);

                int firstDifatSector = fatCount;
                int directoryStart = fatCount + difatCount;
                int miniFatStart = directoryStart + directorySectors;
                int miniStreamStart = miniFatStart + miniFatSectors;
                int bigStart = miniStreamStart + miniStreamSectors;

                layout.Header = BuildHeader(
                    fatCount, difatCount, difatCount > 0 ? firstDifatSector : EndOfChain, directoryStart, directorySectors,
                    miniFatSectors > 0 ? miniFatStart : EndOfChain, miniFatSectors);
                layout.Fat = BuildFat(fatCount, difatCount, directoryStart, directorySectors, miniFatStart, miniFatSectors,
                    miniStreamStart, miniStreamSectors, bigStart, bigSectorCount);
                layout.Difat = difatCount > 0 ? BuildDifat(fatCount, difatCount, firstDifatSector) : [];
                layout.MiniFat = miniFatSectors == 0 ? [] : BuildMiniFat(miniFatSectors, mini.Count, firstMiniSector, streams, mini);
                layout.Directory = BuildDirectory(
                    streams, directorySectors,
                    miniStreamSectors > 0 ? miniStreamStart : EndOfChain, miniBlobBytes,
                    mini, firstMiniSector, big, bigStart, bigSectorCount);
                return layout;
            }

            private static (int fatCount, int difatCount) ComputeSectorCounts(int payloadSectors)
            {
                int fat = CeilingDiv(payloadSectors, FatEntriesPerSector);
                for (int i = 0; i < 4; i++)
                {
                    int difat = fat <= MaxHeaderDifat ? 0 : CeilingDiv(fat - MaxHeaderDifat, DifatEntriesPerSector);
                    fat = CeilingDiv(fat + difat + payloadSectors, FatEntriesPerSector);
                }
                int finalDifat = fat <= MaxHeaderDifat ? 0 : CeilingDiv(fat - MaxHeaderDifat, DifatEntriesPerSector);
                return (fat, finalDifat);
            }

            private static byte[] BuildHeader(int fatCount, int difatCount, int firstDifatSector, int directoryStart,
                int directorySectors, int firstMiniFatSector, int miniFatSectors)
            {
                byte[] header = new byte[HeaderSize];
                ReadOnlySpan<byte> signature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
                signature.CopyTo(header);
                WriteU16(header, 0x18, 0x003E);
                WriteU16(header, 0x1A, 0x0003);
                WriteU16(header, 0x1C, 0xFFFE);
                WriteU16(header, 0x1E, 9);
                WriteU16(header, 0x20, 6);
                _ = directorySectors;
                WriteI32(header, 0x2C, fatCount);
                WriteI32(header, 0x30, directoryStart);
                WriteI32(header, 0x38, MiniCutoff);
                WriteI32(header, 0x3C, firstMiniFatSector);
                WriteI32(header, 0x40, miniFatSectors);
                WriteI32(header, 0x44, firstDifatSector);
                WriteI32(header, 0x48, difatCount);

                for (int i = 0x4C; i < HeaderSize; i += 4)
                {
                    WriteI32(header, i, FreeSector);
                }
                int headerFatCount = Math.Min(fatCount, MaxHeaderDifat);
                for (int i = 0; i < headerFatCount; i++)
                {
                    WriteI32(header, 0x4C + (i * 4), i);
                }
                return header;
            }

            private static byte[] BuildFat(int fatCount, int difatCount, int directoryStart, int directorySectors,
                int miniFatStart, int miniFatSectors, int miniStreamStart, int miniStreamSectors,
                int bigStart, int[] bigSectorCount)
            {
                byte[] fat = new byte[fatCount * FatEntriesPerSector * 4];
                fat.AsSpan().Fill(0xFF);

                for (int i = 0; i < fatCount; i++)
                {
                    WriteI32(fat, i * 4, FatSectorMarker);
                }
                for (int i = 0; i < difatCount; i++)
                {
                    WriteI32(fat, (fatCount + i) * 4, DifatSectorMarker);
                }

                ChainRun(fat, directoryStart, directorySectors);
                ChainRun(fat, miniFatStart, miniFatSectors);
                ChainRun(fat, miniStreamStart, miniStreamSectors);
                int cursor = bigStart;
                foreach (int sectors in bigSectorCount)
                {
                    ChainRun(fat, cursor, sectors);
                    cursor += sectors;
                }
                return fat;
            }

            private static void ChainRun(byte[] table, int start, int count)
            {
                for (int i = 0; i < count; i++)
                {
                    int sector = start + i;
                    WriteI32(table, sector * 4, i == count - 1 ? EndOfChain : sector + 1);
                }
            }

            private static byte[] BuildMiniFat(int miniFatSectors, int miniCount, int[] firstMiniSector,
                IReadOnlyList<CfbStreamSpec> streams, List<int> mini)
            {
                byte[] miniFat = new byte[miniFatSectors * FatEntriesPerSector * 4];
                miniFat.AsSpan().Fill(0xFF);
                for (int k = 0; k < miniCount; k++)
                {
                    int sectors = CeilingDivLong(streams[mini[k]].Size, MiniSectorSize);
                    ChainRun(miniFat, firstMiniSector[k], sectors);
                }
                return miniFat;
            }

            private static byte[] BuildDirectory(IReadOnlyList<CfbStreamSpec> streams, int directorySectors,
                int miniStreamStart, long miniBlobBytes, List<int> mini, int[] firstMiniSector,
                List<int> big, int bigStart, int[] bigSectorCount)
            {
                byte[] directory = new byte[directorySectors * SectorSize];
                for (int i = streams.Count + 1; i < directorySectors * DirectoryEntriesPerSector; i++)
                {
                    Span<byte> unused = directory.AsSpan(i * DirectoryEntrySize, DirectoryEntrySize);
                    WriteI32(unused, 68, NoStream);
                    WriteI32(unused, 72, NoStream);
                    WriteI32(unused, 76, NoStream);
                }

                int[] sorted = [.. Enumerable.Range(0, streams.Count).OrderBy(i => streams[i].Name.Length)
                    .ThenBy(i => streams[i].Name, StringComparer.OrdinalIgnoreCase)];

                WriteEntry(directory.AsSpan(0, DirectoryEntrySize), "Root Entry", objectType: 5,
                    startSector: miniStreamStart, size: miniBlobBytes, child: sorted[0] + 1,
                    leftSibling: NoStream, rightSibling: NoStream);

                for (int rank = 0; rank < sorted.Length; rank++)
                {
                    int index = sorted[rank];
                    int startSector;
                    int miniPos = mini.IndexOf(index);
                    if (miniPos >= 0)
                    {
                        startSector = firstMiniSector[miniPos];
                    }
                    else
                    {
                        int bigPos = big.IndexOf(index);
                        startSector = bigStart;
                        for (int k = 0; k < bigPos; k++)
                        {
                            startSector += bigSectorCount[k];
                        }
                    }
                    int right = rank + 1 < sorted.Length ? sorted[rank + 1] + 1 : NoStream;
                    WriteEntry(directory.AsSpan((index + 1) * DirectoryEntrySize, DirectoryEntrySize),
                        streams[index].Name, objectType: 2, startSector: startSector, size: streams[index].Size,
                        child: NoStream, leftSibling: NoStream, rightSibling: right);
                }
                return directory;
            }

            private static void WriteEntry(Span<byte> entry, string name, byte objectType, int startSector,
                long size, int child, int leftSibling, int rightSibling)
            {
                if (name.Length > 31)
                {
                    throw new ArgumentException($"CFB stream name '{name}' exceeds 31 characters.", nameof(name));
                }
                System.Text.Encoding.Unicode.GetBytes(name + '\0').CopyTo(entry);
                WriteU16(entry, 64, (ushort)((name.Length + 1) * 2));
                entry[66] = objectType;
                entry[67] = 1;
                WriteI32(entry, 68, leftSibling);
                WriteI32(entry, 72, rightSibling);
                WriteI32(entry, 76, child);
                WriteI32(entry, 116, startSector);
                BinaryPrimitives.WriteInt64LittleEndian(entry[120..], size);
            }

            private static byte[] BuildDifat(int fatCount, int difatCount, int firstDifatSector)
            {
                byte[] difat = new byte[difatCount * SectorSize];
                for (int d = 0; d < difatCount; d++)
                {
                    int byteBase = d * SectorSize;
                    int fatBase = MaxHeaderDifat + (d * DifatEntriesPerSector);
                    for (int j = 0; j < DifatEntriesPerSector; j++)
                    {
                        int fatIdx = fatBase + j;
                        WriteI32(difat, byteBase + (j * 4), fatIdx < fatCount ? fatIdx : FreeSector);
                    }
                    int next = (d < difatCount - 1) ? firstDifatSector + d + 1 : EndOfChain;
                    WriteI32(difat, byteBase + (DifatEntriesPerSector * 4), next);
                }
                return difat;
            }
        }

        internal static ValueTask WriteAsync(Stream destination, int workbookSize, Func<Stream, CancellationToken, ValueTask> writeBody, CancellationToken ct)
        {
            int storedSize = Math.Max(CeilingDiv(workbookSize, SectorSize) * SectorSize, MiniCutoff);
            int padding = storedSize - workbookSize;
            CfbStreamSpec spec = new(
                "Workbook",
                storedSize,
                WriteBody: static _ => throw new NotSupportedException("This spec carries an async body only."),
                WriteBodyAsync: async (stream, token) =>
                {
                    await writeBody(stream, token).ConfigureAwait(false);
                    await WritePaddingAsync(stream, padding, token).ConfigureAwait(false);
                });
            return WriteAsync(destination, [spec], ct);
        }

        internal static void Write(Stream destination, int workbookSize, Action<Stream> writeBody)
        {
            int storedSize = Math.Max(CeilingDiv(workbookSize, SectorSize) * SectorSize, MiniCutoff);
            int padding = storedSize - workbookSize;
            CfbStreamSpec spec = new(
                "Workbook",
                storedSize,
                WriteBody: stream =>
                {
                    writeBody(stream);
                    WritePadding(stream, padding);
                },
                WriteBodyAsync: static (_, _) => throw new NotSupportedException("This spec carries a sync body only."));
            Write(destination, [spec]);
        }
    }
}
