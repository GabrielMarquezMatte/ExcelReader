namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Excludes a property from both parsing and record writing: no column is read into it and none is
    /// written from it. Use for computed or transient properties.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public sealed class ExcelIgnoreAttribute : Attribute
    {
    }
}
