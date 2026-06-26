namespace ExcelReader.Core.Parser
{
    // Marks a property whose column must be present in the header row (parsing throws when the header
    // is read if no name matches). By default the cell must also be non-empty in every data row, with
    // a per-row throw on the first blank. Set AllowEmpty = true to require only the column's presence.
    // Presence/non-empty only — it does not validate that the value actually parses.
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public sealed class ExcelRequiredAttribute : Attribute
    {
        public bool AllowEmpty { get; set; }
    }
}
