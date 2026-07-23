namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Marks a property whose column must be present in the header row; parsing throws when the header
    /// is read if no matching name is found. By default every data row's cell must also be non-empty
    /// (and must parse successfully — a non-empty but unparsable cell fails the same way as a blank one),
    /// throwing on the first row that fails either check. Set <see cref="AllowEmpty"/> to true to require
    /// only the column's presence. Contrast with <see cref="ExcelParserConfig.ThrowOnParseFailure"/>,
    /// which governs parse-failure behavior only for non-required columns.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public sealed class ExcelRequiredAttribute : Attribute
    {
        /// <summary>
        /// When true, only the column's presence in the header is required; individual rows may leave
        /// the cell blank or unparsable without throwing. Defaults to false.
        /// </summary>
        public bool AllowEmpty { get; set; }
    }
}
