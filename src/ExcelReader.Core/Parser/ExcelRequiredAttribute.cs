namespace ExcelReader.Core.Parser
{
    // Marks a property whose column must be present in the header row (parsing throws when the header
    // is read if no name matches). By default the cell must also be non-empty in every data row, with
    // a per-row throw on the first blank. Set AllowEmpty = true to require only the column's presence.
    // A cell that is non-empty but fails to parse into the property's type is treated the same as a
    // blank cell and throws the same "missing required value" error — presence alone is not enough,
    // the value must actually parse (see ExcelParserConfig.ThrowOnParseFailure for non-required columns).
    /// <summary>
    /// Marks a property whose column must be present in the header row; parsing throws when the header
    /// is read if no matching name is found. By default every data row's cell must also be non-empty
    /// (and must parse successfully), throwing on the first row that fails either check. Set
    /// <see cref="AllowEmpty"/> to true to require only the column's presence.
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
