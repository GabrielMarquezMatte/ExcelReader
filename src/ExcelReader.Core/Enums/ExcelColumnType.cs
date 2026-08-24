namespace ExcelReader.Core.Enums
{
    /// <summary>
    /// A column's inferred or declared value type, as produced by
    /// <see cref="Reader.Excel.InferSchema(Reader.IExcelRowReader, int, int)"/>.
    /// </summary>
    /// <remarks>
    /// The underlying values are fixed: they are the XL_T_* constants of the native C ABI
    /// (see <c>src/ExcelReader.Native/include/excelreader.h</c>), which marshals this enum with a
    /// plain integer cast rather than a translation table. Renumbering a member is a silent ABI break.
    /// </remarks>
    public enum ExcelColumnType
    {
        /// <summary>UTF-8 text. Also the fallback for any column whose sampled cells disagreed.</summary>
        StringColumn = 0,
        /// <summary>A 64-bit signed integer.</summary>
        Int64Column = 1,
        /// <summary>A 64-bit IEEE 754 floating-point number.</summary>
        Float64Column = 2,
        /// <summary>A boolean.</summary>
        BoolColumn = 3,
        /// <summary>A whole date, counted in days since 1970-01-01.</summary>
        DateColumn = 4,
        /// <summary>A time of day, counted in microseconds since midnight.</summary>
        TimeColumn = 5,
        /// <summary>A date and time, counted in microseconds since 1970-01-01T00:00:00Z.</summary>
        TimestampColumn = 6,
    }
}