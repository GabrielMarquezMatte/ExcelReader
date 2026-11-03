namespace ExcelReader.NetFramework
{
    /// <summary>One column of an inferred schema.</summary>
    public sealed class ExcelColumnSchema
    {
        internal ExcelColumnSchema(string? name, ExcelColumnType type, bool nullable)
        {
            Name = name;
            Type = type;
            Nullable = nullable;
        }

        /// <summary>The header text, or null when the sheet has no header row.</summary>
        public string? Name { get; }

        /// <summary>The inferred type.</summary>
        public ExcelColumnType Type { get; }

        /// <summary>True when the sampled rows had empty cells in the column.</summary>
        public bool Nullable { get; }
    }
}
