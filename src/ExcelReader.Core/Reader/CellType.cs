using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// Identifies the kind of value a <see cref="Cell"/> holds.
    /// </summary>
    /// <remarks>A formula cell reports the type of its cached result, the same in every format.</remarks>
    [SuppressMessage("Design", "CA1027:Mark enums with FlagsAttribute", Justification = "Not flags: the values match the native XL_CELL_* codes, and 5 belonged to the removed Formula.")]
    public enum CellType
    {
        /// <summary>The cell has no value.</summary>
        Empty,
        /// <summary>The cell holds text (a shared string, an inline string, or a formula's text result).</summary>
        ExcelString,
        /// <summary>The cell holds a numeric value.</summary>
        Number,
        /// <summary>The cell holds a numeric value whose style indicates a date/time format.</summary>
        Date,
        /// <summary>The cell holds a boolean value.</summary>
        Boolean,
        /// <summary>The cell holds an Excel error value (e.g. #DIV/0!).</summary>
        Error = 6,
    }
}
