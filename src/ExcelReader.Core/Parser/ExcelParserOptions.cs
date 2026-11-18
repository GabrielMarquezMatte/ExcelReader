using System.Globalization;

namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Options controlling how <see cref="ExcelParser{T}"/> locates the header row, matches header
    /// text to bound properties, and handles cells that fail to parse.
    /// </summary>
    public sealed class ExcelParserOptions
    {
        /// <summary>Gets the 1-based row number that contains column headers. Defaults to <c>1</c>.</summary>
        public int HeaderRow { get; init; } = 1;

        /// <summary>Gets the comparer used to match header text to the names bound by property attributes. Defaults to <see cref="StringComparer.OrdinalIgnoreCase"/>.</summary>
        public StringComparer ColumnNameComparer { get; init; } = StringComparer.OrdinalIgnoreCase;

        /// <summary>Gets how header text is normalized before it is compared against bound property names. Defaults to <see cref="HeaderNormalization.Trim"/>.</summary>
        public HeaderNormalization HeaderNormalization { get; init; } = HeaderNormalization.Trim;

        /// <summary>
        /// Gets the culture used when parsing text cells (CSV fields, XLSX/XLSB/XLS strings) as numbers or
        /// <see cref="Guid"/> (e.g. pt-BR "1.234,56"). Numeric and date cells stored by the workbook are
        /// always read invariantly and ignore this setting.
        /// Defaults to <see cref="CultureInfo.InvariantCulture"/>.
        /// </summary>
        public CultureInfo Culture { get; init; } = CultureInfo.InvariantCulture;

        /// <summary>
        /// Gets whether a non-empty cell that fails to parse into its bound property's type throws
        /// <c>ExcelParseException</c> instead of silently leaving the property at its default. Defaults to
        /// <see langword="false"/>. Independent of <c>[ExcelRequired]</c>, which always throws on a failed
        /// parse of a required column regardless of this flag.
        /// </summary>
        public bool ThrowOnParseFailure { get; init; }
    }
}
