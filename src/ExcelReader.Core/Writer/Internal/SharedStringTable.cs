using System.Globalization;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Writer.Xlsb;

namespace ExcelReader.Core.Writer.Internal
{
    /// <summary>
    /// The workbook's shared strings. A single writer's texts live in one dictionary, looked up with one
    /// hash; parallel workers read it without a lock and add their new texts to shards split by hash,
    /// so workers adding different texts take different locks. The <c>GetOrAdd</c> overloads never run
    /// while workers call <see cref="GetOrAddConcurrent"/>, which is what keeps those lock-free reads safe.
    /// </summary>
    internal sealed class SharedStringTable
    {
        private const int ShardCount = 64;

        private readonly Shard[] _shards = new Shard[ShardCount];
        private readonly Dictionary<string, int> _serial = new(StringComparer.Ordinal);
        private readonly Dictionary<string, int>.AlternateLookup<ReadOnlySpan<char>> _serialSpan;
        private int _count;
        private int _uniqueCount;
        private bool _sharded;

        internal SharedStringTable()
        {
            for (int i = 0; i < ShardCount; i++)
            {
                _shards[i] = new Shard();
            }
            _serialSpan = _serial.GetAlternateLookup<ReadOnlySpan<char>>();
        }

        internal int Count => Volatile.Read(ref _count);

        internal int UniqueCount => Volatile.Read(ref _uniqueCount);

        internal int GetOrAdd(string value)
        {
            _count++;
            ref int index = ref CollectionsMarshal.GetValueRefOrAddDefault(_serial, value, out bool exists);
            if (!exists)
            {
                index = SerialMiss(value);
            }
            return index;
        }

        internal int GetOrAdd(ReadOnlySpan<char> value)
        {
            _count++;
            ref int index = ref CollectionsMarshal.GetValueRefOrAddDefault(_serialSpan, value, out bool exists);
            if (!exists)
            {
                index = SerialMiss(value);
            }
            return index;
        }

        // Workers may have added it to a shard; a text new to both takes the next index.
        private int SerialMiss(ReadOnlySpan<char> value)
        {
            return _sharded && ShardFor(value).SpanIndexes.TryGetValue(value, out int index) ? index : _uniqueCount++;
        }

        /// <summary>Resolves a text from a parallel worker without counting the reference; see <see cref="AddReferences"/>.</summary>
        internal int GetOrAddConcurrent(ReadOnlySpan<char> value)
        {
            if (_serialSpan.TryGetValue(value, out int serial))
            {
                return serial;
            }
            Shard shard = ShardFor(value);
            lock (shard.Gate)
            {
                if (shard.SpanIndexes.TryGetValue(value, out int index))
                {
                    return index;
                }
                index = Interlocked.Increment(ref _uniqueCount) - 1;
                shard.SpanIndexes.TryAdd(value, index);
                _sharded = true;
                return index;
            }
        }

        /// <summary>Counts references resolved without the counting <c>GetOrAdd</c> overloads.</summary>
        internal void AddReferences(int count)
        {
            Interlocked.Add(ref _count, count);
        }

        // The part is written to its entry in pieces of about this size, not built whole first.
        internal const int FlushBytes = 64 * 1024;

        internal void WriteTo(Stream entry, bool binary)
        {
            string[] values = Values();
            using BiffBuffer buffer = new(FlushBytes + 4096);
            using BiffBuffer payload = new(128);
            int next = WriteStart(buffer, payload, binary);
            while (next <= values.Length)
            {
                next = Fill(buffer, payload, values, next, binary);
                entry.Write(buffer.Span);
                buffer.Reset();
            }
        }

        internal async ValueTask WriteToAsync(Stream entry, bool binary, CancellationToken ct)
        {
            string[] values = Values();
            using BiffBuffer buffer = new(FlushBytes + 4096);
            using BiffBuffer payload = new(128);
            int next = WriteStart(buffer, payload, binary);
            while (next <= values.Length)
            {
                next = Fill(buffer, payload, values, next, binary);
                await entry.WriteAsync(buffer.Memory, ct).ConfigureAwait(false);
                buffer.Reset();
            }
        }

        private int WriteStart(BiffBuffer buffer, BiffBuffer payload, bool binary)
        {
            if (binary)
            {
                payload.WriteU32((uint)Count);
                payload.WriteU32((uint)UniqueCount);
                Biff12RecordWriter.WriteRecord(buffer, Brt.BeginSst, payload.Span);
                return 0;
            }
            buffer.WriteUtf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<sst xmlns=\"{XlsxConstants.MainNs}\" count=\"");
            buffer.WriteUtf8(Count.ToString(CultureInfo.InvariantCulture));
            buffer.Write("\" uniqueCount=\""u8);
            buffer.WriteUtf8(UniqueCount.ToString(CultureInfo.InvariantCulture));
            buffer.Write("\">"u8);
            return 0;
        }

        // Appends items from next until the buffer reaches FlushBytes, and the closing record after the
        // last one; returns where to resume, past values.Length once the part is complete.
        private static int Fill(BiffBuffer buffer, BiffBuffer payload, string[] values, int next, bool binary)
        {
            while (next < values.Length && buffer.Length < FlushBytes)
            {
                WriteItem(buffer, payload, values[next++], binary);
            }
            if (next < values.Length)
            {
                return next;
            }
            if (binary)
            {
                Biff12RecordWriter.WriteRecord(buffer, Brt.EndSst);
                return next + 1;
            }
            buffer.Write("</sst>"u8);
            return next + 1;
        }

        private static void WriteItem(BiffBuffer buffer, BiffBuffer payload, string value, bool binary)
        {
            if (binary)
            {
                payload.Reset();
                payload.WriteByte(0);
                Biff12RecordWriter.WriteWideString(payload, value);
                Biff12RecordWriter.WriteRecord(buffer, Brt.SSTItem, payload.Span);
                return;
            }
            buffer.Write("<si><t"u8);
            if (NeedsPreserveSpace(value))
            {
                buffer.Write(" xml:space=\"preserve\""u8);
            }
            buffer.WriteByte((byte)'>');
            CellFormatter.WriteEscaped(buffer, value);
            buffer.Write("</t></si>"u8);
        }

        private string[] Values()
        {
            string[] values = new string[UniqueCount];
            foreach (KeyValuePair<string, int> entry in _serial)
            {
                values[entry.Value] = entry.Key;
            }
            foreach (Shard shard in _shards)
            {
                foreach (KeyValuePair<string, int> entry in shard.Indexes)
                {
                    values[entry.Value] = entry.Key;
                }
            }
            return values;
        }

        private Shard ShardFor(ReadOnlySpan<char> value)
        {
            return _shards[string.GetHashCode(value, StringComparison.Ordinal) & (ShardCount - 1)];
        }

        private static bool NeedsPreserveSpace(string value)
        {
            return value.Length > 0
                && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]));
        }

        private sealed class Shard
        {
            internal Shard()
            {
                SpanIndexes = Indexes.GetAlternateLookup<ReadOnlySpan<char>>();
            }

            internal Dictionary<string, int> Indexes { get; } = new(StringComparer.Ordinal);

            internal Dictionary<string, int>.AlternateLookup<ReadOnlySpan<char>> SpanIndexes { get; }

            internal Lock Gate { get; } = new();
        }
    }
}
