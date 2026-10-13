using System.Buffers.Binary;
using System.Runtime.InteropServices;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Xls
{
    /// <summary>
    /// The SST and EXTSST records for text longer than <see cref="MaxLabelLength"/>: Excel truncates a
    /// LABEL cell to that length on load and refuses a LABEL continued past one record, so longer text
    /// goes through LABELSST instead.
    /// </summary>
    /// <remarks>
    /// Each string is encoded once, straight into its SST form (character count, flags, characters),
    /// and deduplicated by comparing those bytes; the record boundaries are only drawn in
    /// <see cref="Write"/>.
    /// </remarks>
    internal sealed class XlsSharedStrings : IDisposable
    {
        internal const int MaxLabelLength = 255;

        private const int Continue = 0x003C;
        private const int Sst = 0x00FC;
        private const int ExtSst = 0x00FF;
        private const int EntryHeaderLength = 3;

        // Strings are found by a hash of a few sampled bytes. Text that differs only between the samples
        // piles into one chain; once a chain holds this many, its strings move to a hash of every byte,
        // so no chain is ever searched linearly past this length.
        private const int MaxSampleChain = 8;
        private const int Crowded = -1;

        private BiffBuffer? _entries;
        private readonly List<int> _starts = [];
        private readonly List<int> _next = [];
        private readonly Dictionary<int, int> _bySample = [];
        private readonly Dictionary<int, int> _byFullHash = [];
        private int _references;

        internal int Add(ReadOnlySpan<char> value)
        {
            _references++;
            BiffBuffer entries = _entries ??= new BiffBuffer(64 * 1024);
            int start = entries.Length;
            bool compressed = BiffStringEncoder.CanCompress(value);
            entries.WriteU16(value.Length);
            entries.WriteByte((byte)(compressed ? 0 : 1));
            BiffStringEncoder.WriteChars(entries, value, compressed);
            ReadOnlySpan<byte> entry = entries.Span[start..];

            ref int head = ref CollectionsMarshal.GetValueRefOrAddDefault(_bySample, SampleHash(entry), out bool exists);
            if (!exists)
            {
                head = Append(start, next: -1);
                return head;
            }
            if (head != Crowded)
            {
                int length = 0;
                for (int candidate = head; candidate >= 0; candidate = _next[candidate], length++)
                {
                    if (Entry(candidate, start).SequenceEqual(entry))
                    {
                        entries.Truncate(start);
                        return candidate;
                    }
                }
                if (length < MaxSampleChain)
                {
                    head = Append(start, next: head);
                    return head;
                }
                MoveToFullHash(head, start);
                head = Crowded;
            }

            ref int fullHead = ref CollectionsMarshal.GetValueRefOrAddDefault(_byFullHash, FullHash(entry), out bool hashed);
            if (hashed)
            {
                for (int candidate = fullHead; candidate >= 0; candidate = _next[candidate])
                {
                    if (Entry(candidate, start).SequenceEqual(entry))
                    {
                        entries.Truncate(start);
                        return candidate;
                    }
                }
            }
            fullHead = Append(start, next: hashed ? fullHead : -1);
            return fullHead;
        }

        private int Append(int start, int next)
        {
            _starts.Add(start);
            _next.Add(next);
            return _starts.Count - 1;
        }

        private void MoveToFullHash(int chain, int end)
        {
            for (int index = chain; index >= 0;)
            {
                int following = _next[index];
                ref int fullHead = ref CollectionsMarshal.GetValueRefOrAddDefault(_byFullHash, FullHash(Entry(index, end)), out bool hashed);
                _next[index] = hashed ? fullHead : -1;
                fullHead = index;
                index = following;
            }
        }

        /// <summary>An upper bound on what <see cref="Write"/> adds, so the globals buffer can be sized once.</summary>
        internal int EstimatedLength
        {
            get
            {
                if (_entries is null)
                {
                    return 0;
                }
                int records = (_entries.Length / (BiffRecord.MaxPayload - EntryHeaderLength)) + 2;
                return _entries.Length + (records * 5) + 12 + ((_starts.Count / 8) + 1) * 8;
            }
        }

        /// <summary>Writes SST and EXTSST at the buffer's end, which must be their position in the Workbook stream.</summary>
        internal void Write(BiffBuffer buffer)
        {
            int count = _starts.Count;
            if (_entries is null || count == 0)
            {
                return;
            }
            int perBucket = Math.Max(8, (count + 127) / 128);
            var buckets = new List<(int Position, int RecordOffset)>((count + perBucket - 1) / perBucket);

            int lengthPos = buffer.BeginRecord(Sst);
            buffer.WriteI32(_references);
            buffer.WriteI32(count);
            for (int i = 0; i < count; i++)
            {
                ReadOnlySpan<byte> entry = Entry(i, _entries.Length);
                byte flags = entry[2];
                int charSize = flags == 0 ? 1 : 2;
                if (Room(buffer, lengthPos) < EntryHeaderLength + charSize)
                {
                    lengthPos = NextRecord(buffer, lengthPos);
                }
                if (i % perBucket == 0)
                {
                    buckets.Add((buffer.Length, buffer.Length - (lengthPos - 2)));
                }
                buffer.Write(entry[..EntryHeaderLength]);
                ReadOnlySpan<byte> chars = entry[EntryHeaderLength..];
                while (!chars.IsEmpty)
                {
                    int take = Math.Min(Room(buffer, lengthPos) / charSize * charSize, chars.Length);
                    if (take == 0)
                    {
                        lengthPos = NextRecord(buffer, lengthPos);
                        buffer.WriteByte(flags);
                        continue;
                    }
                    buffer.Write(chars[..take]);
                    chars = chars[take..];
                }
            }
            buffer.EndRecord(lengthPos);

            int extSst = buffer.BeginRecord(ExtSst);
            buffer.WriteU16(perBucket);
            foreach ((int position, int recordOffset) in buckets)
            {
                buffer.WriteI32(position);
                buffer.WriteU16(recordOffset);
                buffer.WriteU16(0);
            }
            buffer.EndRecord(extSst);
        }

        public void Dispose()
        {
            _entries?.Dispose();
        }

        private ReadOnlySpan<byte> Entry(int index, int end)
        {
            int start = _starts[index];
            int next = index + 1 < _starts.Count ? _starts[index + 1] : end;
            return _entries!.Span[start..next];
        }

        /// <summary>
        /// Hashes the length, the first and last 32 bytes and three 8-byte samples in between: hashing
        /// all of every long string cost more than the rest of writing it. Every entry is longer than
        /// 255 bytes, so each read is in range.
        /// </summary>
        private static int SampleHash(ReadOnlySpan<byte> entry)
        {
            ReadOnlySpan<ulong> head = MemoryMarshal.Cast<byte, ulong>(entry[..32]);
            ReadOnlySpan<ulong> tail = MemoryMarshal.Cast<byte, ulong>(entry[^32..]);
            var hasher = default(HashCode);
            hasher.Add(entry.Length);
            hasher.Add(head[0] ^ tail[0]);
            hasher.Add(head[1] ^ tail[1]);
            hasher.Add(head[2] ^ tail[2]);
            hasher.Add(head[3] ^ tail[3]);
            hasher.Add(BinaryPrimitives.ReadUInt64LittleEndian(entry[(entry.Length / 4)..]));
            hasher.Add(BinaryPrimitives.ReadUInt64LittleEndian(entry[(entry.Length / 2)..]));
            hasher.Add(BinaryPrimitives.ReadUInt64LittleEndian(entry[(entry.Length * 3 / 4)..]));
            return hasher.ToHashCode();
        }

        private static int FullHash(ReadOnlySpan<byte> entry)
        {
            var hasher = default(HashCode);
            hasher.AddBytes(entry);
            return hasher.ToHashCode();
        }

        private static int Room(BiffBuffer buffer, int lengthPos)
        {
            return BiffRecord.MaxPayload - (buffer.Length - lengthPos - 2);
        }

        private static int NextRecord(BiffBuffer buffer, int lengthPos)
        {
            buffer.EndRecord(lengthPos);
            return buffer.BeginRecord(Continue);
        }
    }
}
