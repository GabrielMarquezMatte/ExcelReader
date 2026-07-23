using System.Globalization;

namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Options controlling how <see cref="ExcelParser{T}"/> and <c>RefParser</c> locate the header
    /// row, match header text to bound properties, and handle cells that fail to parse.
    /// </summary>
    public sealed class ExcelParserConfig
    {
        /// <summary>Gets the 1-based row number that contains column headers. Defaults to <c>1</c>.</summary>
        public int HeaderRow { get; init; } = 1;

        /// <summary>Gets the comparer used to match header text to the names bound by property attributes. Defaults to <see cref="StringComparer.OrdinalIgnoreCase"/>.</summary>
        public StringComparer ColumnNameComparer { get; init; } = StringComparer.OrdinalIgnoreCase;

        /// <summary>Gets how header text is normalized before it is compared against bound property names. Defaults to <see cref="Parser.HeaderNormalization.Trim"/>.</summary>
        public HeaderNormalization HeaderNormalization { get; init; } = HeaderNormalization.Trim;

        // Culture used when parsing text-backed numeric and Guid cells (e.g. pt-BR "1.234,56").
        // Binary numeric cells (XLS/XLSB) carry a raw double and ignore this. Defaults to invariant
        // to preserve existing behavior.
        /// <summary>
        /// Gets the culture used when parsing text-backed numeric and <see cref="Guid"/> cells (e.g. pt-BR
        /// "1.234,56"). Binary numeric cells from XLS/XLSB carry a raw double and ignore this setting.
        /// Defaults to <see cref="CultureInfo.InvariantCulture"/>.
        /// </summary>
        public CultureInfo Culture { get; init; } = CultureInfo.InvariantCulture;

        // When true, a non-empty cell that fails to parse into its bound property's type throws
        // ExcelParseException instead of silently leaving the property at its default. Defaults to
        // false to preserve existing lenient behavior. Independent of [ExcelRequired]: a required
        // column whose cell fails to parse always throws (as "missing required value"), regardless
        // of this flag, since the property never actually received a value either way.
        /// <summary>
        /// Gets whether a non-empty cell that fails to parse into its bound property's type throws
        /// <c>ExcelParseException</c> instead of silently leaving the property at its default. Defaults to
        /// <see langword="false"/>. Independent of <c>[ExcelRequired]</c>, which always throws on a failed
        /// parse of a required column regardless of this flag.
        /// </summary>
        public bool ThrowOnParseFailure { get; init; }
    }
}
