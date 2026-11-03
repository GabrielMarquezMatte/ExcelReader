using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.NetFramework
{
    /// <summary>The kind of value a raw cell holds.</summary>
    public enum ExcelCellType
    {
        /// <summary>No value.</summary>
        Empty = 0,

        /// <summary>A string.</summary>
        [SuppressMessage("Naming", "CA1720:Identifier contains type name", Justification = "Public name of the cell kind.")]
        String = 1,

        /// <summary>A number.</summary>
        Number = 2,

        /// <summary>
        /// A date or time. The cell text is the Excel serial number in the workbook's date system
        /// (see <see cref="ExcelWorkbook.IsDate1904"/>).
        /// </summary>
        Date = 3,

        /// <summary>A boolean.</summary>
        Bool = 4,

        /// <summary>A formula result.</summary>
        Formula = 5,

        /// <summary>An error value.</summary>
        Error = 6,
    }
}
