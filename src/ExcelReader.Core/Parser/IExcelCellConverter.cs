using ExcelReader.Core.Reader;


namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Converts a matched cell into a property value. Implement this and attach the implementation with
    /// <c>[ExcelConverter(typeof(MyConverter))]</c> for types the built-in parsers do not handle (money
    /// strings, custom date formats, domain value objects, ...).
    /// </summary>
    /// <remarks>
    /// <typeparamref name="T"/> must be the exact property type (use
    /// <see cref="IExcelCellConverter{T}"/> of <c>decimal?</c> for a <c>decimal?</c> property). A single
    /// instance is created once and shared across every row and thread, so implementations must be
    /// stateless and thread-safe.
    /// </remarks>
    public interface IExcelCellConverter<T>
    {
        /// <summary>
        /// Converts a non-empty cell into a value of type <typeparamref name="T"/>; return false to
        /// signal a parse failure, which leaves the target property at its default.
        /// </summary>
        /// <param name="cell">The non-empty cell to convert.</param>
        /// <param name="context">The date system and format provider to convert with.</param>
        /// <param name="value">The converted value, when this method returns true.</param>
        bool TryConvert(in Cell cell, in ExcelRowContext context, out T value);
    }
}
