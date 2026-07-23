using System.Globalization;

namespace ExcelReader.Core.Parser
{
    // Thrown when a bound column's cell is non-empty but fails to parse into the target property type,
    // and ExcelParserConfig.ThrowOnParseFailure is true. With the default (false), a parse failure on a
    // non-required column silently keeps the property at its default (existing, pre-F3 behavior); a
    // parse failure on an [ExcelRequired] column instead surfaces as ProjectionRules.MissingRequiredValue
    // regardless of this flag, since the column effectively has no usable value either way.
    public sealed class ExcelParseException : Exception
    {
        public ExcelParseException()
        {
        }

        public ExcelParseException(string message)
            : base(message)
        {
        }

        public ExcelParseException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        public ExcelParseException(int row, string columnName, string rawValue)
            : base(string.Create(CultureInfo.InvariantCulture,
                $"Failed to parse column '{columnName}' in row {row}: '{rawValue}'."))
        {
            Row = row;
            ColumnName = columnName;
            RawValue = rawValue;
        }

        public int Row { get; }
        public string ColumnName { get; } = string.Empty;
        public string RawValue { get; } = string.Empty;
    }
}
