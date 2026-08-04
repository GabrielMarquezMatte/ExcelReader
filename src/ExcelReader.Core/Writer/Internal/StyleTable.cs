using System.Runtime.InteropServices;

namespace ExcelReader.Core.Writer.Internal
{
    // Deduplicated, ordered registry of CellStyle values shared by the four workbook writers'
    // AddStyle implementations. Index 0 is always the general/default style and index 1 is always
    // the builtin date style — every date cell already renders through that fixed index today, so
    // this class seeds both unconditionally rather than waiting for a caller to register them.
    internal sealed class StyleTable
    {
        // Matches the numFmtId=14 "mm-dd-yy" builtin format every writer already hardcodes for dates.
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

        // Assigns a custom numFmtId (>=164, the first id below that range being reserved for Excel's
        // builtins) to every distinct NumberFormat string among the custom styles (index 2+). Shared
        // by the XLSX/XLSB/XLS style serializers so the "custom ids start at 164" rule (R-B3) and the
        // "one id per distinct format string" dedup live in exactly one place.
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

        // Assigns a font index to every distinct (Bold, Italic) combination used by a registered
        // style. (false, false) always maps to font index 0 — the single default font every writer
        // already emits — so a workbook with no bold/italic styles adds no extra font record.
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
