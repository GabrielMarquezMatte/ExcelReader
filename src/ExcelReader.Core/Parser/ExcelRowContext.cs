using System.Globalization;

namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Settings a cell reader, converter or row parser needs to convert a cell, handed to every
    /// <see cref="ExcelCellReader{TValue}"/>, <see cref="ExcelRowParser{TModel}"/> and
    /// <see cref="IExcelCellConverter{T}.TryConvert"/> call.
    /// </summary>
    /// <remarks>
    /// Passed by <see langword="in"/> reference. New settings arrive as new properties, so a converter
    /// written against this type keeps compiling. Construct one with an object initializer to unit-test a
    /// converter; <see langword="default"/> means the 1900 date system and the invariant culture.
    /// </remarks>
    public readonly record struct ExcelRowContext
    {
        private readonly IFormatProvider? _formatProvider;

        /// <summary>True when the source workbook uses the 1904 date system.</summary>
        public bool IsDate1904 { get; init; }

        /// <summary>The format provider to use when parsing cell text. Defaults to <see cref="CultureInfo.InvariantCulture"/>.</summary>
        public IFormatProvider FormatProvider
        {
            get => _formatProvider ?? CultureInfo.InvariantCulture;
            init => _formatProvider = value;
        }
    }
}
