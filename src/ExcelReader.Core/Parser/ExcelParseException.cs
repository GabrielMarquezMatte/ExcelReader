using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Thrown when a bound column's cell is non-empty but fails to parse into the target property type
    /// and the parser is configured to throw on parse failures (rather than silently keep the property
    /// at its default). See <see cref="ExcelParserConfig.ThrowOnParseFailure"/>.
    /// </summary>
    /// <remarks>
    /// A parse failure on an <c>[ExcelRequired]</c> column always surfaces as a missing-required-value
    /// error instead, regardless of <see cref="ExcelParserConfig.ThrowOnParseFailure"/>, since the column
    /// has no usable value either way.
    /// </remarks>
    public sealed class ExcelParseException : Exception
    {
        /// <summary>Creates an exception with no message.</summary>
        [ExcludeFromCodeCoverage]
        public ExcelParseException()
        {
        }

        /// <summary>Creates an exception with the given message.</summary>
        [ExcludeFromCodeCoverage]
        public ExcelParseException(string message)
            : base(message)
        {
        }

        /// <summary>Creates an exception with the given message and inner exception.</summary>
        [ExcludeFromCodeCoverage]
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

        /// <summary>Creates an exception describing a required column that has no value in a given row.</summary>
        public ExcelParseException(int row, string columnName)
            : base(string.Create(CultureInfo.InvariantCulture,
                $"Required column '{columnName}' has no value in row {row}."))
        {
            Row = row;
            ColumnName = columnName;
        }

        /// <summary>Creates an exception describing required column(s) missing from the header row.</summary>
        public ExcelParseException(IReadOnlyList<string> missingColumnNames)
            : base($"Required column(s) not found in the header row: {string.Join(", ", missingColumnNames)}.")
        {
            ColumnName = string.Join(", ", missingColumnNames);
        }

        /// <summary>The 1-based row number where the parse failure occurred (the header row is row 1).</summary>
        public int Row { get; }
        /// <summary>The name of the column whose value failed to parse.</summary>
        public string ColumnName { get; } = string.Empty;
        /// <summary>The raw cell text that could not be parsed.</summary>
        public string RawValue { get; } = string.Empty;
    }
}
