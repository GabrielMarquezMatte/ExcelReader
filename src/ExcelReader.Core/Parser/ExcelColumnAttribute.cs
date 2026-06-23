namespace ExcelReader.Core.Parser
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public sealed class ExcelColumnAttribute : Attribute
    {
        public string Name { get; }

        public ExcelColumnAttribute(string name)
        {
            ArgumentNullException.ThrowIfNull(name);
            Name = name;
        }
    }
}
