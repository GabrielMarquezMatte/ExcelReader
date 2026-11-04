using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Writer.Internal
{
    internal sealed class StyleTable
    {
        internal static readonly CellStyle DateStyle = new() { NumberFormat = "mm-dd-yy" };

        // Excel refuses a workbook with more custom number formats than this (its documented limit
        // is "between 200 and 250").
        private const int MaxCustomNumberFormats = 200;

        private readonly Dictionary<CellStyle, int> _indices = [];
        private readonly List<CellStyle> _styles = [];
        private readonly HashSet<string> _customFormats = new(StringComparer.Ordinal);
        private readonly Lock _dateVariantGate = new();
        private int[] _dateVariants = [];

        private readonly StyleLimits _limits;
        private readonly HashSet<FontKey> _fonts = [default];
        private readonly List<ExcelColor> _palette = [];
        private readonly HashSet<ExcelColor> _paletteSet = [];

        internal StyleTable(StyleLimits limits)
        {
            _limits = limits;
            _styles.Add(default);
            _indices[default] = 0;
            _styles.Add(DateStyle);
            _indices[DateStyle] = 1;
        }

        internal IReadOnlyList<CellStyle> Styles => _styles;

        internal IReadOnlyList<ExcelColor> PaletteColors => _palette;

        internal int Count => _styles.Count;

        internal int Add(CellStyle style)
        {
            if (_indices.TryGetValue(style, out int index))
            {
                return index;
            }
            CellStyleValidation.Validate(style, nameof(style));
            CheckLimits(style, out bool newFont, out List<ExcelColor>? newColors);
            ClaimNumberFormat(style.NumberFormat);
            if (newFont)
            {
                _fonts.Add(FontKey.Of(style));
            }
            foreach (ExcelColor color in newColors ?? [])
            {
                _palette.Add(color);
                _paletteSet.Add(color);
            }
            index = _styles.Count;
            _styles.Add(style);
            _indices[style] = index;
            return index;
        }

        // Parallel record writers resolve this per date cell on every worker, so a known variant is read
        // without the lock; entries hold variant + 1, leaving 0 for "not resolved yet".
        internal int DateVariant(int styleId)
        {
            int[] known = Volatile.Read(ref _dateVariants);
            if ((uint)styleId < (uint)known.Length && known[styleId] != 0)
            {
                return known[styleId] - 1;
            }
            return AddDateVariant(styleId);
        }

        private int AddDateVariant(int styleId)
        {
            lock (_dateVariantGate)
            {
                int[] known = _dateVariants;
                if ((uint)styleId < (uint)known.Length && known[styleId] != 0)
                {
                    return known[styleId] - 1;
                }
                CellStyle style = _styles[styleId];
                int variant = style.NumberFormat is null ? Add(style with { NumberFormat = DateStyle.NumberFormat }) : styleId;
                int[] next = new int[Math.Max(known.Length, styleId + 1)];
                known.CopyTo(next, 0);
                next[styleId] = variant + 1;
                Volatile.Write(ref _dateVariants, next);
                return variant;
            }
        }

        // Runs once per overlay push, before any worker starts; the gate keeps it safe next to DateVariant's registrations.
        internal int InheritNumberFormat(int overlayStyleId, int columnStyleId)
        {
            if (columnStyleId == 0)
            {
                return overlayStyleId;
            }
            lock (_dateVariantGate)
            {
                CellStyle overlay = _styles[overlayStyleId];
                string? columnFormat = _styles[columnStyleId].NumberFormat;
                if (overlay.NumberFormat is not null || columnFormat is null)
                {
                    return overlayStyleId;
                }
                return Add(overlay with { NumberFormat = columnFormat });
            }
        }

        private void ClaimNumberFormat(string? format)
        {
            if (format is null || _customFormats.Contains(format))
            {
                return;
            }
            if (_customFormats.Count == MaxCustomNumberFormats)
            {
                throw new ExcelLimitExceededException("NumberFormats", MaxCustomNumberFormats, MaxCustomNumberFormats + 1L);
            }
            _customFormats.Add(format);
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

        internal StyleComponents AssignComponents()
        {
            var fonts = new List<FontKey> { default };
            var fills = new List<FillKey> { default, default };
            var borders = new List<BorderKey> { default };
            var fontIndex = new Dictionary<FontKey, int> { [default] = 0 };
            var fillIndex = new Dictionary<FillKey, int> { [default] = 0 };
            var borderIndex = new Dictionary<BorderKey, int> { [default] = 0 };
            int[] fontIds = new int[_styles.Count];
            int[] fillIds = new int[_styles.Count];
            int[] borderIds = new int[_styles.Count];
            for (int i = 0; i < _styles.Count; i++)
            {
                CellStyle style = _styles[i];
                fontIds[i] = IdOf(FontKey.Of(style), fonts, fontIndex);
                fillIds[i] = style.FillColor is null ? 0 : IdOf(new FillKey(style.FillColor), fills, fillIndex);
                borderIds[i] = IdOf(BorderKey.Of(style), borders, borderIndex);
            }
            return new StyleComponents(fonts, fills, borders, fontIds, fillIds, borderIds);
        }

        private static int IdOf<TKey>(TKey key, List<TKey> list, Dictionary<TKey, int> index)
            where TKey : notnull
        {
            if (index.TryGetValue(key, out int id))
            {
                return id;
            }
            id = list.Count;
            list.Add(key);
            index[key] = id;
            return id;
        }

        private void CheckLimits(in CellStyle style, out bool newFont, out List<ExcelColor>? newColors)
        {
            if (_styles.Count + 1 > _limits.MaxCellXfs)
            {
                throw new ExcelLimitExceededException("CellStyles", _limits.MaxCellXfs, _styles.Count + 1L);
            }
            newFont = !_fonts.Contains(FontKey.Of(style));
            if (newFont && _fonts.Count + 1 + _limits.ReservedFontIndices > _limits.MaxFonts)
            {
                throw new ExcelLimitExceededException("Fonts", _limits.MaxFonts, _fonts.Count + 1L + _limits.ReservedFontIndices);
            }
            newColors = null;
            CollectNewColor(style.FontColor, ref newColors);
            CollectNewColor(style.FillColor, ref newColors);
            CollectNewColor(style.Left?.Color, ref newColors);
            CollectNewColor(style.Right?.Color, ref newColors);
            CollectNewColor(style.Top?.Color, ref newColors);
            CollectNewColor(style.Bottom?.Color, ref newColors);
            if (newColors is not null && _palette.Count + newColors.Count > _limits.MaxPaletteColors)
            {
                throw new ExcelLimitExceededException("Colors", _limits.MaxPaletteColors, _palette.Count + (long)newColors.Count);
            }
        }

        private void CollectNewColor(ExcelColor? color, ref List<ExcelColor>? fresh)
        {
            if (color is not { } value || _paletteSet.Contains(value) || fresh?.Contains(value) == true)
            {
                return;
            }
            fresh ??= [];
            fresh.Add(value);
        }
    }
}
