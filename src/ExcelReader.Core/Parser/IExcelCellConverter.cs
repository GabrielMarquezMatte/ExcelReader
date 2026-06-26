using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser
{
    // Converts a matched cell into a property value. Attach to a property with
    // [ExcelConverter(typeof(MyConverter))] for types the built-in parsers do not handle
    // (money strings, custom date formats, domain value objects, ...).
    //
    // T must be the exact property type (use IExcelCellConverter<decimal?> for a decimal? property).
    // A single instance is created once and shared across every row and thread, so implementations
    // must be stateless / thread-safe. Empty cells are skipped before TryConvert runs, so the
    // property keeps its default; return false to signal a parse failure (also keeps the default).
    public interface IExcelCellConverter<T>
    {
        bool TryConvert(in Cell cell, bool isDate1904, IFormatProvider provider, out T value);
    }
}
