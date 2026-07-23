namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Maps a property to a worksheet column by header name. Apply more than once to accept alternate
    /// header spellings for the same property.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
    public sealed class ExcelColumnAttribute : Attribute
    {
        /// <summary>The header name this attribute matches.</summary>
        public string Name { get; }

        /// <summary>Creates an attribute that maps the property to the column with the given header name.</summary>
        public ExcelColumnAttribute(string name)
        {
            ArgumentNullException.ThrowIfNull(name);
            Name = name;
        }
    }
}
