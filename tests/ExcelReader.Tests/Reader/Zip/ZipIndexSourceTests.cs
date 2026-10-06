using System.IO.Compression;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Sources;
using ExcelReader.Core.Reader.Zip;

namespace ExcelReader.Tests.Reader.Zip
{
    public sealed class ZipIndexSourceTests : IDisposable
    {
        public enum Kind
        {
            Memory,
            OpaqueMemoryStream,
            Trickle,
            File,
        }

        private static readonly string[] Names = ["deflated.txt", "stored.bin", "empty"];
        private static readonly int[] Sizes = [300_000, 70_000, 0];

        private readonly List<string> _paths = [];

        public void Dispose()
        {
            foreach (string path in _paths)
            {
                File.Delete(path);
            }
        }

        public static TheoryData<Kind, bool> KindsAndDescriptors
        {
            get
            {
                TheoryData<Kind, bool> data = [];
                foreach (Kind kind in Enum.GetValues<Kind>())
                {
                    data.Add(kind, false);
                    data.Add(kind, true);
                }
                return data;
            }
        }

        public static TheoryData<Kind> Kinds => [Kind.Memory, Kind.OpaqueMemoryStream, Kind.Trickle, Kind.File];

        private static byte[] Content(int entry, int scale = 1)
        {
            byte[] bytes = new byte[Sizes[entry] / scale];
            for (int i = 0; i < bytes.Length; i++)
            {
                bytes[i] = (byte)((i % 251) ^ (entry * 37));
            }
            return bytes;
        }

        private sealed class ForwardOnlyStream(Stream inner) : Stream
        {
            public override bool CanRead => false;

            public override bool CanSeek => false;

            public override bool CanWrite => true;

            public override long Length => throw new NotSupportedException();

            public override long Position
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public override void Flush()
            {
                inner.Flush();
            }

            public override int Read(byte[] buffer, int offset, int count)
            {
                throw new NotSupportedException();
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
                inner.Write(buffer, offset, count);
            }
        }

        private static byte[] BuildZip(bool dataDescriptors, int scale = 1)
        {
            using MemoryStream buffer = new();
            using ForwardOnlyStream forward = new(buffer);
            using (ZipArchive zip = new(dataDescriptors ? forward : buffer, ZipArchiveMode.Create, leaveOpen: true))
            {
                for (int i = 0; i < Names.Length; i++)
                {
                    ZipArchiveEntry entry = zip.CreateEntry(Names[i], i == 1 ? CompressionLevel.NoCompression : CompressionLevel.Optimal);
                    using Stream stream = entry.Open();
                    stream.Write(Content(i, scale));
                }
            }
            return buffer.ToArray();
        }

        private ByteSource Open(Kind kind, byte[] zipBytes)
        {
            switch (kind)
            {
                case Kind.Memory:
                    return ByteSource.FromMemory(zipBytes);
                case Kind.OpaqueMemoryStream:
                    return ByteSource.FromStream(new MemoryStream(zipBytes, writable: false), leaveOpen: false);
                case Kind.Trickle:
                    return ByteSource.FromStream(new TrickleStream(zipBytes), leaveOpen: false);
                default:
                    string path = Path.Combine(Path.GetTempPath(), "excelreader-zipindex-" + Guid.NewGuid().ToString("N") + ".zip");
                    File.WriteAllBytes(path, zipBytes);
                    _paths.Add(path);
                    return ByteSource.OpenFile(path, asynchronous: false);
            }
        }

        private static DecompressedByteCounter Counter()
        {
            return new DecompressedByteCounter(ExcelReaderOptions.Default.MaxTotalDecompressedBytes);
        }

        [Theory]
        [MemberData(nameof(KindsAndDescriptors))]
        public void Parts_Hold_The_Entry_Bytes(Kind kind, bool dataDescriptors)
        {
            using ZipIndex index = ZipIndex.Create(Open(kind, BuildZip(dataDescriptors)), ExcelReaderOptions.Default);
            Assert.Equal(kind == Kind.Memory, index.HasMemory);
            for (int i = 0; i < Names.Length; i++)
            {
                Assert.True(index.TryGetEntry(System.Text.Encoding.UTF8.GetBytes(Names[i]), out ZipEntryRef entry));
                using ZipPart part = index.OpenPart(entry, Counter());
                Assert.True(part.Memory.Span.SequenceEqual(Content(i)));
            }
        }

        [Theory]
        [MemberData(nameof(KindsAndDescriptors))]
        public void Entry_Streams_Yield_The_Entry_Bytes(Kind kind, bool dataDescriptors)
        {
            using ZipIndex index = ZipIndex.Create(Open(kind, BuildZip(dataDescriptors)), ExcelReaderOptions.Default);
            for (int i = 0; i < Names.Length; i++)
            {
                Assert.True(index.TryGetEntry(System.Text.Encoding.UTF8.GetBytes(Names[i]), out ZipEntryRef entry));
                using Stream stream = index.OpenEntryStream(entry, Counter(), ExcelReaderOptions.Default);
                using MemoryStream copy = new();
                stream.CopyTo(copy);
                Assert.True(copy.ToArray().AsSpan().SequenceEqual(Content(i)));
            }
        }

        [Theory]
        [MemberData(nameof(KindsAndDescriptors))]
        public async Task Async_Open_Matches_Sync_Open(Kind kind, bool dataDescriptors)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using ZipIndex index = await ZipIndex.CreateAsync(Open(kind, BuildZip(dataDescriptors)), ExcelReaderOptions.Default, ct);
            for (int i = 0; i < Names.Length; i++)
            {
                byte[] name = System.Text.Encoding.UTF8.GetBytes(Names[i]);
                using ZipPart part = await index.OpenPartOrDefaultAsync(name, Counter(), ct);
                Assert.True(part.Memory.Span.SequenceEqual(Content(i)));

                Assert.True(index.TryGetEntry(name, out ZipEntryRef entry));
                await using Stream stream = await index.OpenEntryStreamAsync(entry, Counter(), ExcelReaderOptions.Default, ct);
                using MemoryStream copy = new();
                await stream.CopyToAsync(copy, ct);
                Assert.True(copy.ToArray().AsSpan().SequenceEqual(Content(i)));
            }
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void Entries_Open_Concurrently(Kind kind)
        {
            using ZipIndex index = ZipIndex.Create(Open(kind, BuildZip(dataDescriptors: false, scale: kind == Kind.Trickle ? 20 : 1)), ExcelReaderOptions.Default);
            int scale = kind == Kind.Trickle ? 20 : 1;
            Parallel.For(0, 24, worker =>
            {
                int which = worker % 2;
                Assert.True(index.TryGetEntry(System.Text.Encoding.UTF8.GetBytes(Names[which]), out ZipEntryRef entry));
                using Stream stream = index.OpenEntryStream(entry, Counter(), ExcelReaderOptions.Default);
                using MemoryStream copy = new();
                stream.CopyTo(copy);
                Assert.True(copy.ToArray().AsSpan().SequenceEqual(Content(which, scale)));
            });
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void A_File_Cut_Short_Is_Invalid_Data(Kind kind)
        {
            byte[] zipBytes = BuildZip(dataDescriptors: false, scale: 20);
            Assert.Throws<InvalidDataException>(
                () => ZipIndex.Create(Open(kind, zipBytes[..(zipBytes.Length / 2)]), ExcelReaderOptions.Default));
            Assert.Throws<InvalidDataException>(
                () => ZipIndex.Create(Open(kind, []), ExcelReaderOptions.Default));
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void A_Stored_Entry_Whose_Sizes_Disagree_Is_Invalid_Data(Kind kind)
        {
            byte[] zipBytes = BuildZip(dataDescriptors: false, scale: 20);
            byte[] name = System.Text.Encoding.UTF8.GetBytes(Names[1]);
            int record = -1;
            for (int i = 0; i + 46 + name.Length <= zipBytes.Length; i++)
            {
                if (zipBytes.AsSpan(i, 4).SequenceEqual<byte>([0x50, 0x4B, 0x01, 0x02]) && zipBytes.AsSpan(i + 46, name.Length).SequenceEqual(name))
                {
                    record = i;
                    break;
                }
            }
            Assert.True(record >= 0, "central directory record not found");
            zipBytes.AsSpan(record + 24, 4).Clear();

            using ZipIndex index = ZipIndex.Create(Open(kind, zipBytes), ExcelReaderOptions.Default);
            Assert.True(index.TryGetEntry(name, out ZipEntryRef entry));
            Assert.Throws<InvalidDataException>(() => index.OpenPart(entry, Counter()).Dispose());
            Assert.Throws<InvalidDataException>(() => index.OpenEntryStream(entry, Counter(), ExcelReaderOptions.Default).Dispose());
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void Bytes_After_The_End_Of_Central_Directory_Are_Ignored(Kind kind)
        {
            byte[] zipBytes = BuildZip(dataDescriptors: false);
            foreach (int extra in (int[])[1, 100])
            {
                byte[] padded = new byte[zipBytes.Length + extra];
                zipBytes.CopyTo(padded, 0);
                padded.AsSpan(zipBytes.Length).Fill(0xAB);

                using ZipIndex index = ZipIndex.Create(Open(kind, padded), ExcelReaderOptions.Default);
                for (int i = 0; i < Names.Length; i++)
                {
                    Assert.True(index.TryGetEntry(System.Text.Encoding.UTF8.GetBytes(Names[i]), out ZipEntryRef entry));
                    using ZipPart part = index.OpenPart(entry, Counter());
                    Assert.True(part.Memory.Span.SequenceEqual(Content(i)));
                }
            }
        }

        [Fact]
        public void A_Failed_Create_Disposes_The_Source()
        {
            TrickleStream stream = new(new byte[100]);
            Assert.Throws<InvalidDataException>(
                () => ZipIndex.Create(ByteSource.FromStream(stream, leaveOpen: false), ExcelReaderOptions.Default));
            Assert.False(stream.CanRead);
        }

        private enum Access
        {
            Part,
            Stream,
            AsyncPart,
            AsyncStream,
        }

        private static async Task<byte[]> ReadEntryAsync(ZipIndex index, ZipEntryRef entry, Access access, CancellationToken ct)
        {
            switch (access)
            {
                case Access.Part:
                    using (ZipPart part = index.OpenPart(entry, Counter()))
                    {
                        return part.Memory.ToArray();
                    }
                case Access.AsyncPart:
                    using (ZipPart part = await index.OpenPartAsync(entry, Counter(), ct))
                    {
                        return part.Memory.ToArray();
                    }
                case Access.Stream:
                    using (Stream stream = index.OpenEntryStream(entry, Counter(), ExcelReaderOptions.Default))
                    {
                        using MemoryStream copy = new();
                        stream.CopyTo(copy);
                        return copy.ToArray();
                    }
                default:
                    await using (Stream stream = await index.OpenEntryStreamAsync(entry, Counter(), ExcelReaderOptions.Default, ct))
                    {
                        using MemoryStream copy = new();
                        await stream.CopyToAsync(copy, ct);
                        return copy.ToArray();
                    }
            }
        }

        private static async Task<string> OutcomeAsync(Func<ByteSource> open, Access access, CancellationToken ct)
        {
            try
            {
                using ZipIndex index = access is Access.AsyncPart or Access.AsyncStream
                    ? await ZipIndex.CreateAsync(open(), ExcelReaderOptions.Default, ct)
                    : ZipIndex.Create(open(), ExcelReaderOptions.Default);
                System.Text.StringBuilder text = new();
                foreach (string name in Names)
                {
                    if (!index.TryGetEntry(System.Text.Encoding.UTF8.GetBytes(name), out ZipEntryRef entry))
                    {
                        text.Append("missing;");
                        continue;
                    }
                    try
                    {
                        byte[] bytes = await ReadEntryAsync(index, entry, access, ct);
                        text.Append(Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes))).Append(';');
                    }
                    catch (Exception ex) when (ex is InvalidDataException or NotSupportedException or ExcelLimitExceededException)
                    {
                        text.Append(ex.GetType().Name).Append(';');
                    }
                }
                return text.ToString();
            }
            catch (Exception ex) when (ex is InvalidDataException or NotSupportedException or ExcelLimitExceededException)
            {
                return ex.GetType().Name;
            }
        }

        [Fact]
        public async Task Mutated_Bytes_Give_The_Same_Outcome_From_Memory_And_From_A_Stream()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            byte[] pristine = BuildZip(dataDescriptors: false, scale: 1000);
            uint state = 2024;
            for (int i = 0; i < 3000; i++)
            {
                byte[] mutated = (byte[])pristine.Clone();
                state = (state * 1664525) + 1013904223;
                int position = (int)(state % (uint)mutated.Length);
                state = (state * 1664525) + 1013904223;
                mutated[position] = (byte)(state >> 24);

                foreach (Access access in Enum.GetValues<Access>())
                {
                    Access reference = access is Access.Part or Access.AsyncPart ? Access.Part : Access.Stream;
                    string fromMemory = await OutcomeAsync(() => ByteSource.FromMemory(mutated), reference, ct);
                    string fromStream = await OutcomeAsync(
                        () => ByteSource.FromStream(new MemoryStream(mutated, writable: false), leaveOpen: false), access, ct);
                    Assert.True(string.Equals(fromMemory, fromStream, StringComparison.Ordinal), $"mutation {i} at byte {position}, {access}: memory={fromMemory} stream={fromStream}");
                }
            }
        }

        private static int CentralRecordOffset(byte[] zipBytes, string entryName)
        {
            byte[] name = System.Text.Encoding.UTF8.GetBytes(entryName);
            for (int i = 0; i + 46 + name.Length <= zipBytes.Length; i++)
            {
                if (zipBytes.AsSpan(i, 4).SequenceEqual<byte>([0x50, 0x4B, 0x01, 0x02]) && zipBytes.AsSpan(i + 46, name.Length).SequenceEqual(name))
                {
                    return i;
                }
            }
            throw new InvalidOperationException("central directory record not found");
        }

        private static byte[] ZipDeclaringTooMuchData()
        {
            byte[] zipBytes = BuildZip(dataDescriptors: false, scale: 20);
            int record = CentralRecordOffset(zipBytes, Names[0]);
            System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(zipBytes.AsSpan(record + 24), (uint)(Sizes[0] / 20) + 64);
            return zipBytes;
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void A_Deflated_Entry_Shorter_Than_Declared_Is_Invalid_Data(Kind kind)
        {
            using ZipIndex index = ZipIndex.Create(Open(kind, ZipDeclaringTooMuchData()), ExcelReaderOptions.Default);
            Assert.True(index.TryGetEntry(System.Text.Encoding.UTF8.GetBytes(Names[0]), out ZipEntryRef entry));
            InvalidDataException ex = Assert.Throws<InvalidDataException>(() => index.OpenPart(entry, Counter()).Dispose());
            Assert.Contains("less data", ex.Message, StringComparison.Ordinal);
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public async Task A_Deflated_Entry_Shorter_Than_Declared_Is_Invalid_Data_Asynchronously(Kind kind)
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            using ZipIndex index = await ZipIndex.CreateAsync(Open(kind, ZipDeclaringTooMuchData()), ExcelReaderOptions.Default, ct);
            Assert.True(index.TryGetEntry(System.Text.Encoding.UTF8.GetBytes(Names[0]), out ZipEntryRef entry));
            InvalidDataException ex = await Assert.ThrowsAsync<InvalidDataException>(
                async () => (await index.OpenPartAsync(entry, Counter(), ct)).Dispose());
            Assert.Contains("less data", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_Workbooks_Match_ZipArchive_Entry_For_Entry()
        {
            string dataDirectory = Path.Combine(AppContext.BaseDirectory, "data");
            int checkedFiles = 0;
            foreach (string path in Directory.EnumerateFiles(dataDirectory, "*.xls?", SearchOption.AllDirectories))
            {
                byte[] bytes = File.ReadAllBytes(path);
                if (bytes.Length < 4 || bytes[0] != (byte)'P' || bytes[1] != (byte)'K')
                {
                    continue;
                }
                using ZipArchive expected = new(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
                using ZipIndex index = ZipIndex.Create(ByteSource.OpenFile(path, asynchronous: false), ExcelReaderOptions.Default);
                foreach (ZipArchiveEntry expectedEntry in expected.Entries)
                {
                    if (expectedEntry.FullName.EndsWith('/'))
                    {
                        continue;
                    }
                    Assert.True(index.TryGetEntry(System.Text.Encoding.UTF8.GetBytes(expectedEntry.FullName), out ZipEntryRef entry), expectedEntry.FullName);
                    using Stream expectedStream = expectedEntry.Open();
                    using MemoryStream expectedBytes = new();
                    expectedStream.CopyTo(expectedBytes);
                    using ZipPart part = index.OpenPart(entry, Counter());
                    Assert.True(part.Memory.Span.SequenceEqual(expectedBytes.ToArray()), path + "!" + expectedEntry.FullName);
                }
                checkedFiles++;
            }
            Assert.True(checkedFiles > 0, "no ZIP-based workbook found under " + dataDirectory);
        }

        [Theory]
        [MemberData(nameof(Kinds))]
        public void A_Disposed_Index_Refuses_Every_Open(Kind kind)
        {
            ZipIndex index = ZipIndex.Create(Open(kind, BuildZip(dataDescriptors: false, scale: 20)), ExcelReaderOptions.Default);
            Assert.True(index.TryGetEntry(System.Text.Encoding.UTF8.GetBytes(Names[0]), out ZipEntryRef entry));
            index.Dispose();

            Assert.Throws<ObjectDisposedException>(() => index.TryGetEntry(System.Text.Encoding.UTF8.GetBytes(Names[0]), out _));
            Assert.Throws<ObjectDisposedException>(() => index.OpenPart(entry, Counter()));
            Assert.Throws<ObjectDisposedException>(() => index.OpenEntryStream(entry, Counter(), ExcelReaderOptions.Default));
        }

        [Fact]
        public async Task A_Failed_CreateAsync_Disposes_The_Source()
        {
            TrickleStream stream = new(new byte[100]);
            await Assert.ThrowsAsync<InvalidDataException>(async () => await ZipIndex.CreateAsync(
                ByteSource.FromStream(stream, leaveOpen: false), ExcelReaderOptions.Default, TestContext.Current.CancellationToken));
            Assert.False(stream.CanRead);
        }

        [Fact]
        public async Task A_Cancelled_CreateAsync_Disposes_The_Source()
        {
            TrickleStream stream = new(BuildZip(dataDescriptors: false, scale: 20));
            using CancellationTokenSource cancelled = new();
            await cancelled.CancelAsync();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await ZipIndex.CreateAsync(
                ByteSource.FromStream(stream, leaveOpen: false), ExcelReaderOptions.Default, cancelled.Token));
            Assert.False(stream.CanRead);
        }
    }
}
