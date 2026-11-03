namespace ExcelReader.NetFramework.Parser
{
    /// <summary>
    /// Excludes a property from parsing: no column is read into it. Use for computed or transient properties.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public sealed class ExcelIgnoreAttribute : Attribute
    {
    }
}
