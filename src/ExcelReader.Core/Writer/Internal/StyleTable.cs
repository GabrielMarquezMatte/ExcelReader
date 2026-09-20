using System.Runtime.InteropServices;

namespace ExcelReader.Core.Writer.Internal
{
    internal sealed class StyleTable
    {
        internal static readonly CellStyle DateStyle = new() { NumberFormat = "mm-dd-yy" };

        private readonly Dictionary<CellStyle, int> _indices = [];
        private readonly List<CellStyle> _styles = [];

        internal StyleTable()
        {
            _styles.Add(default);
            _indices[default] = 0;
            _styles.Add(DateStyle);
            _indices[DateStyle] = 1;
        }

        internal IReadOnlyList<CellStyle> Styles => _styles;

        internal int Count => _styles.Count;

        internal int Add(CellStyle style)
        {
            if (_indices.TryGetValue(style, out int index))
            {
                return index;
            }
            index = _styles.Count;
            _styles.Add(style);
            _indices[style] = index;
            return index;
        }

        internal Dictionary<string, int> AssignCustomNumberFormatIds()
        {
            var ids = new Dictionary<string, int>(StringComparer.Ordinal);
            int next = 164;
            for (int i = 2; i < _styles.Count; i++)
            {
                string? format = _styles[i].NumberFormat;
                if (format is not null && !ids.ContainsKey(format))
                {
                    ids[format] = next;
                    next++;
                }
            }
            return ids;
        }

        internal Dictionary<(bool Bold, bool Italic), int> AssignFontIds()
        {
            var ids = new Dictionary<(bool Bold, bool Italic), int> { [(false, false)] = 0 };
            int next = 1;
            foreach (ref readonly CellStyle style in CollectionsMarshal.AsSpan(_styles))
            {
                (bool, bool) key = (style.Bold, style.Italic);
                if (!ids.ContainsKey(key))
                {
                    ids[key] = next;
                    next++;
                }
            }
            return ids;
        }
    }
}
