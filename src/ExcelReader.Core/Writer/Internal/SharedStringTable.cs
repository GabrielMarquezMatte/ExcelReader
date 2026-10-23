using System.Globalization;
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
            if (_serial.TryGetValue(value, out int index))
            {
                return index;
            }
            index = SerialMiss(value);
            _serial.Add(value, index);
            return index;
        }

        internal int GetOrAdd(ReadOnlySpan<char> value)
        {
            _count++;
            if (_serialSpan.TryGetValue(value, out int index))
            {
                return index;
            }
            index = SerialMiss(value);
            _serialSpan.TryAdd(value, index);
            return index;
        }

        // Workers may have added it to a shard; a text new to both takes the next index.
        private int SerialMiss(ReadOnlySpan<char> value)
        {
            return ShardFor(value).SpanIndexes.TryGetValue(value, out int index) ? index : _uniqueCount++;
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
                return index;
            }
        }

        /// <summary>Counts references resolved without the counting <c>GetOrAdd</c> overloads.</summary>
        internal void AddReferences(int count)
        {
            Interlocked.Add(ref _count, count);
        }

        internal BiffBuffer ToXlsxBytes()
        {
            string[] values = Values();
            BiffBuffer xml = new(Math.Max(256, values.Length * 32));
            xml.WriteUtf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<sst xmlns=\"{XlsxConstants.MainNs}\" count=\"");
            xml.WriteUtf8(Count.ToString(CultureInfo.InvariantCulture));
            xml.Write("\" uniqueCount=\""u8);
            xml.WriteUtf8(UniqueCount.ToString(CultureInfo.InvariantCulture));
            xml.Write("\">"u8);
            foreach (string value in values)
            {
                xml.Write("<si><t"u8);
                if (NeedsPreserveSpace(value))
                {
                    xml.Write(" xml:space=\"preserve\""u8);
                }
                xml.WriteByte((byte)'>');
                CellFormatter.WriteEscaped(xml, value);
                xml.Write("</t></si>"u8);
            }
            xml.Write("</sst>"u8);
            return xml;
        }

        internal BiffBuffer ToXlsbBytes()
        {
            string[] values = Values();
            var data = new BiffBuffer(Math.Max(128, values.Length * 24));
            using var payload = new BiffBuffer(128);
            payload.WriteU32((uint)Count);
            payload.WriteU32((uint)UniqueCount);
            Biff12RecordWriter.WriteRecord(data, Brt.BeginSst, payload.Span);
            foreach (string value in values)
            {
                payload.Reset();
                payload.WriteByte(0);
                Biff12RecordWriter.WriteWideString(payload, value);
                Biff12RecordWriter.WriteRecord(data, Brt.SSTItem, payload.Span);
            }
            Biff12RecordWriter.WriteRecord(data, Brt.EndSst);
            return data;
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
