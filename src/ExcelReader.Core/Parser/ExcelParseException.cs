using System.Globalization;

namespace ExcelReader.Core.Parser
{
    // Thrown when a bound column's cell is non-empty but fails to parse into the target property type,
    // and ExcelParserConfig.ThrowOnParseFailure is true. With the default (false), a parse failure on a
    // non-required column silently keeps the property at its default (existing, pre-F3 behavior); a
    // parse failure on an [ExcelRequired] column instead surfaces as ProjectionRules.MissingRequiredValue
    // regardless of this flag, since the column effectively has no usable value either way.
    /// <summary>
    /// Thrown when a bound column's cell is non-empty but fails to parse into the target property type
    /// and the parser is configured to throw on parse failures (rather than silently keep the property
    /// at its default).
    /// </summary>
    public sealed class ExcelParseException : Exception
    {
        /// <summary>Creates an exception with no message.</summary>
        public ExcelParseException()
        {
        }

        /// <summary>Creates an exception with the given message.</summary>
        public ExcelParseException(string message)
            : base(message)
        {
        }

        /// <summary>Creates an exception with the given message and inner exception.</summary>
        public ExcelParseException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        /// <summary>Creates an exception describing the row, column, and raw value that failed to parse.</summary>
        public ExcelParseException(int row, string columnName, string rawValue)
            : base(string.Create(CultureInfo.InvariantCulture,
                $"Failed to parse column '{columnName}' in row {row}: '{rawValue}'."))
        {
            Row = row;
            ColumnName = columnName;
            RawValue = rawValue;
        }

        /// <summary>The 1-based row number where the parse failure occurred (the header row is row 1).</summary>
        public int Row { get; }
        /// <summary>The name of the column whose value failed to parse.</summary>
        public string ColumnName { get; } = string.Empty;
        /// <summary>The raw cell text that could not be parsed.</summary>
        public string RawValue { get; } = string.Empty;
    }
}
