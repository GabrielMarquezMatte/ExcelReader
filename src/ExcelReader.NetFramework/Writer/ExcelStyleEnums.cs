using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.NetFramework.Writer
{
    /// <summary>The line style of a cell border edge.</summary>
    public enum ExcelBorderStyle
    {
        /// <summary>No line.</summary>
        None = 0,

        /// <summary>A thin line.</summary>
        Thin = 1,

        /// <summary>A medium line.</summary>
        Medium = 2,

        /// <summary>A dashed line.</summary>
        Dashed = 3,

        /// <summary>A dotted line.</summary>
        Dotted = 4,

        /// <summary>A thick line.</summary>
        Thick = 5,

        /// <summary>A double line.</summary>
        [SuppressMessage("Naming", "CA1720:Identifier contains type name", Justification = "The ST_BorderStyle name for a double line.")]
        Double = 6,

        /// <summary>A hairline.</summary>
        Hair = 7,

        /// <summary>A medium dashed line.</summary>
        MediumDashed = 8,

        /// <summary>A dash-dot line.</summary>
        DashDot = 9,

        /// <summary>A medium dash-dot line.</summary>
        MediumDashDot = 10,

        /// <summary>A dash-dot-dot line.</summary>
        DashDotDot = 11,

        /// <summary>A medium dash-dot-dot line.</summary>
        MediumDashDotDot = 12,

        /// <summary>A slanted dash-dot line.</summary>
        SlantDashDot = 13,
    }

    /// <summary>The horizontal alignment of cell content.</summary>
    public enum ExcelHorizontalAlignment
    {
        /// <summary>Text left, numbers right.</summary>
        General = 0,

        /// <summary>Left aligned.</summary>
        Left = 1,

        /// <summary>Centered.</summary>
        Center = 2,

        /// <summary>Right aligned.</summary>
        Right = 3,

        /// <summary>Repeats the content to fill the cell.</summary>
        Fill = 4,

        /// <summary>Justified.</summary>
        Justify = 5,

        /// <summary>Centered across the selection.</summary>
        CenterContinuous = 6,

        /// <summary>Distributed evenly.</summary>
        Distributed = 7,
    }

    /// <summary>The vertical alignment of cell content.</summary>
    public enum ExcelVerticalAlignment
    {
        /// <summary>Bottom aligned.</summary>
        Bottom = 0,

        /// <summary>Top aligned.</summary>
        Top = 1,

        /// <summary>Centered.</summary>
        Center = 2,

        /// <summary>Justified.</summary>
        Justify = 3,

        /// <summary>Distributed evenly.</summary>
        Distributed = 4,
    }
}
