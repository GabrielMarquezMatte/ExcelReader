using System.Globalization;

namespace ExcelReader.Core.Parser
{
    public sealed class ExcelParserConfig
    {
        public int HeaderRow { get; init; } = 1;
        public StringComparer ColumnNameComparer { get; init; } = StringComparer.OrdinalIgnoreCase;
        public HeaderNormalization HeaderNormalization { get; init; } = HeaderNormalization.Trim;

        // Culture used when parsing text-backed numeric and Guid cells (e.g. pt-BR "1.234,56").
        // Binary numeric cells (XLS/XLSB) carry a raw double and ignore this. Defaults to invariant
        // to preserve existing behavior.
        public CultureInfo Culture { get; init; } = CultureInfo.InvariantCulture;

        // When true, a non-empty cell that fails to parse into its bound property's type throws
        // ExcelParseException instead of silently leaving the property at its default. Defaults to
        // false to preserve existing lenient behavior. Independent of [ExcelRequired]: a required
        // column whose cell fails to parse always throws (as "missing required value"), regardless
        // of this flag, since the property never actually received a value either way.
        public bool ThrowOnParseFailure { get; init; }
    }
}
