namespace ExcelReader.NetFramework
{
    /// <summary>A sheet or table of an open workbook.</summary>
    public readonly struct ExcelSheet : IEquatable<ExcelSheet>
    {
        internal ExcelSheet(ExcelWorkbook workbook, int id)
        {
            Workbook = workbook;
            Id = id;
        }

        /// <summary>The workbook this sheet belongs to.</summary>
        public ExcelWorkbook Workbook { get; }

        internal int Id { get; }

        /// <summary>Compares two sheets for equality.</summary>
        /// <param name="left">The first sheet.</param>
        /// <param name="right">The second sheet.</param>
        /// <returns>True when both refer to the same sheet of the same workbook.</returns>
        public static bool operator ==(ExcelSheet left, ExcelSheet right)
        {
            return left.Equals(right);
        }

        /// <summary>Compares two sheets for inequality.</summary>
        /// <param name="left">The first sheet.</param>
        /// <param name="right">The second sheet.</param>
        /// <returns>True when the sheets differ.</returns>
        public static bool operator !=(ExcelSheet left, ExcelSheet right)
        {
            return !left.Equals(right);
        }

        /// <inheritdoc />
        public bool Equals(ExcelSheet other)
        {
            return ReferenceEquals(Workbook, other.Workbook) && Id == other.Id;
        }

        /// <inheritdoc />
        public override bool Equals(object? obj)
        {
            return obj is ExcelSheet other && Equals(other);
        }

        /// <inheritdoc />
        public override int GetHashCode()
        {
            return Id;
        }
    }
}
