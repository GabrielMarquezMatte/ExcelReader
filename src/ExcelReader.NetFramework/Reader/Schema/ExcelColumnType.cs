using System.Diagnostics.CodeAnalysis;
using ExcelReader.NetFramework.Parser.Internal;

namespace ExcelReader.NetFramework.Reader.Schema
{
    /// <summary>The storage type the native reader infers for a column.</summary>
    public enum ExcelColumnType
    {
        /// <summary>Text.</summary>
        [SuppressMessage("Naming", "CA1720:Identifier contains type name", Justification = "Public name of the column type.")]
        String = 0,

        /// <summary>A 64-bit integer.</summary>
        [SuppressMessage("Naming", "CA1720:Identifier contains type name", Justification = "Public name of the column type.")]
        Int64 = 1,

        /// <summary>A double-precision number.</summary>
        [SuppressMessage("Naming", "CA1720:Identifier contains type name", Justification = "Public name of the column type.")]
        Double = 2,

        /// <summary>A boolean.</summary>
        Bool = 3,

        /// <summary>A calendar date.</summary>
        Date = 4,

        /// <summary>A time of day.</summary>
        Time = 5,

        /// <summary>A date with a time of day.</summary>
        Timestamp = 6,
    }
}
