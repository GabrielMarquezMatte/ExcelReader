namespace ExcelReader.NetFramework
{
    /// <summary>A raw cell of a row.</summary>
    public readonly struct ExcelCell : IEquatable<ExcelCell>
    {
        internal ExcelCell(int column, ExcelCellType type, string? text)
        {
            Column = column;
            Type = type;
            Text = text;
        }

        /// <summary>The zero-based column of the cell.</summary>
        public int Column { get; }

        /// <summary>The kind of value the cell holds.</summary>
        public ExcelCellType Type { get; }

        /// <summary>The cell text, or null when the cell is empty.</summary>
        public string? Text { get; }

        /// <summary>Compares two cells for equality.</summary>
        /// <param name="left">The first cell.</param>
        /// <param name="right">The second cell.</param>
        /// <returns>True when column, type and text match.</returns>
        public static bool operator ==(ExcelCell left, ExcelCell right)
        {
            return left.Equals(right);
        }

        /// <summary>Compares two cells for inequality.</summary>
        /// <param name="left">The first cell.</param>
        /// <param name="right">The second cell.</param>
        /// <returns>True when the cells differ.</returns>
        public static bool operator !=(ExcelCell left, ExcelCell right)
        {
            return !left.Equals(right);
        }

        /// <inheritdoc />
        public bool Equals(ExcelCell other)
        {
            return Column == other.Column && Type == other.Type && string.Equals(Text, other.Text, StringComparison.Ordinal);
        }

        /// <inheritdoc />
        public override bool Equals(object? obj)
        {
            return obj is ExcelCell other && Equals(other);
        }

        /// <inheritdoc />
        public override int GetHashCode()
        {
            return Column;
        }
    }
}
