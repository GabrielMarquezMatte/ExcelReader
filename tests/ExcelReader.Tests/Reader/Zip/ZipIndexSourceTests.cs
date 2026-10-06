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

        [Fact]
        public void A_Failed_Create_Disposes_The_Source()
        {
            TrickleStream stream = new(new byte[100]);
            Assert.Throws<InvalidDataException>(
                () => ZipIndex.Create(ByteSource.FromStream(stream, leaveOpen: false), ExcelReaderOptions.Default));
            Assert.False(stream.CanRead);
        }

        private static string Outcome(Func<ByteSource> open)
        {
            try
            {
                using ZipIndex index = ZipIndex.Create(open(), ExcelReaderOptions.Default);
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
                        using ZipPart part = index.OpenPart(entry, Counter());
                        text.Append(Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(part.Memory.Span))).Append(';');
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
        public void Mutated_Bytes_Give_The_Same_Outcome_From_Memory_And_From_A_Stream()
        {
            byte[] pristine = BuildZip(dataDescriptors: false, scale: 1000);
            uint state = 2024;
            for (int i = 0; i < 3000; i++)
            {
                byte[] mutated = (byte[])pristine.Clone();
                state = (state * 1664525) + 1013904223;
                int position = (int)(state % (uint)mutated.Length);
                state = (state * 1664525) + 1013904223;
                mutated[position] = (byte)(state >> 24);

                string fromMemory = Outcome(() => ByteSource.FromMemory(mutated));
                string fromStream = Outcome(() => ByteSource.FromStream(new MemoryStream(mutated, writable: false), leaveOpen: false));
                Assert.True(string.Equals(fromMemory, fromStream, StringComparison.Ordinal),$"mutation {i} at byte {position}: memory={fromMemory} stream={fromStream}");
            }
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
    }
}
