using System.Globalization;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Writer.Internal
{
    internal sealed class SharedStringTable
    {
        private readonly Dictionary<string, int> _indexes = new(StringComparer.Ordinal);
        private readonly List<string> _values = [];

        internal int Count { get; private set; }
        internal int UniqueCount => _values.Count;

        internal int GetOrAdd(string value)
        {
            Count++;
            if (_indexes.TryGetValue(value, out int index))
            {
                return index;
            }
            index = _values.Count;
            _indexes.Add(value, index);
            _values.Add(value);
            return index;
        }

        internal string ToXlsxXml()
        {
            using var xml = new BiffBuffer(Math.Max(256, _values.Count * 32));
            xml.WriteUtf8(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<sst xmlns=\"{XlsxConstants.MainNs}\" count=\"");
            xml.WriteUtf8(Count.ToString(CultureInfo.InvariantCulture));
            xml.Write("\" uniqueCount=\""u8);
            xml.WriteUtf8(UniqueCount.ToString(CultureInfo.InvariantCulture));
            xml.Write("\">"u8);
            foreach (ref readonly string value in CollectionsMarshal.AsSpan(_values))
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
            return System.Text.Encoding.UTF8.GetString(xml.Span);
        }

        internal ReadOnlyMemory<byte> ToXlsbBytes()
        {
            using var data = new BiffBuffer(Math.Max(128, _values.Count * 24));
            using var payload = new BiffBuffer(128);
            foreach (ref readonly string value in CollectionsMarshal.AsSpan(_values))
            {
                payload.Reset();
                payload.WriteByte(0);
                Biff12RecordWriter.WriteWideString(payload, value);
                Biff12RecordWriter.WriteRecord(data, Brt.SSTItem, payload.Span);
            }
            return data.Memory.ToArray();
        }

        private static bool NeedsPreserveSpace(string value)
        {
            return value.Length > 0
                && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]));
        }
    }
}
