using System.Globalization;
using System.Runtime.InteropServices;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Writer.Internal
{
    internal sealed class SharedStringTable
    {
        // Starting from 0 means ~14 rehashes by 50k unique strings, each doubling past ~5,300
        // entries landing on the LOH. A modest starting capacity avoids most of that for typical
        // workbooks while staying cheap for small ones.
        private const int DefaultCapacity = 1024;

        private readonly Dictionary<string, int> _indexes = new(DefaultCapacity, StringComparer.Ordinal);
        private readonly List<string> _values = new(DefaultCapacity);

        internal int Count { get; private set; }
        internal int UniqueCount => _values.Count;

        internal int GetOrAdd(string value)
        {
            Count++;
            ref int index = ref CollectionsMarshal.GetValueRefOrAddDefault(_indexes, value, out bool existed);
            if (existed)
            {
                return index;
            }
            index = _values.Count;
            _values.Add(value);
            return index;
        }

        // Returns the fully-assembled sharedStrings.xml as UTF-8 bytes — no decode-to-string round trip,
        // matching ToXlsbBytes's shape so the caller can write the ZIP entry directly.
        internal BiffBuffer ToXlsxBytes()
        {
            BiffBuffer xml = new(Math.Max(256, _values.Count * 32));
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
            return xml;
        }

        internal BiffBuffer ToXlsbBytes()
        {
            var data = new BiffBuffer(Math.Max(128, _values.Count * 24));
            using var payload = new BiffBuffer(128);
            payload.WriteU32((uint)Count);
            payload.WriteU32((uint)UniqueCount);
            Biff12RecordWriter.WriteRecord(data, Brt.BeginSst, payload.Span);
            foreach (ref readonly string value in CollectionsMarshal.AsSpan(_values))
            {
                payload.Reset();
                payload.WriteByte(0);
                Biff12RecordWriter.WriteWideString(payload, value);
                Biff12RecordWriter.WriteRecord(data, Brt.SSTItem, payload.Span);
            }
            Biff12RecordWriter.WriteRecord(data, Brt.EndSst);
            return data;
        }

        private static bool NeedsPreserveSpace(string value)
        {
            return value.Length > 0
                && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]));
        }
    }
}
