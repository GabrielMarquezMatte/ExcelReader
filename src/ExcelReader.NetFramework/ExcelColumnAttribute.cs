namespace ExcelReader.NetFramework
{
    /// <summary>
    /// Maps a property to a worksheet column by header name. Apply more than once to accept alternate
    /// header spellings for the same property.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
    public sealed class ExcelColumnAttribute : Attribute
    {
        /// <summary>Creates an attribute that maps the property to the column with the given header name.</summary>
        /// <param name="name">The header name.</param>
        public ExcelColumnAttribute(string name)
        {
            if (name is null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            Name = name;
        }

        /// <summary>The header name this attribute matches.</summary>
        public string Name { get; }
    }
}
