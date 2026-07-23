namespace ExcelReader.Core.Enums
{
    /// <summary>
    /// Identifies the kind of value a <see cref="ExcelReader.Core.ValueObjects.Cell"/> holds.
    /// </summary>
    public enum CellType
    {
        /// <summary>The cell has no value.</summary>
        Empty,
        /// <summary>The cell holds text (a shared string or inline string).</summary>
        ExcelString,
        /// <summary>The cell holds a numeric value.</summary>
        Number,
        /// <summary>The cell holds a numeric value whose style indicates a date/time format.</summary>
        Date,
        /// <summary>The cell holds a boolean value.</summary>
        Boolean,
        /// <summary>The cell holds a formula's cached result.</summary>
        Formula,
        /// <summary>The cell holds an Excel error value (e.g. #DIV/0!).</summary>
        Error,
    }
}
