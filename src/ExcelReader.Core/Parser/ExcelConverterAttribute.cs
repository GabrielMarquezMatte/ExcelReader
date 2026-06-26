namespace ExcelReader.Core.Parser
{
    // Binds a custom IExcelCellConverter<TProperty> to a property. The converter type must implement
    // IExcelCellConverter<> for the property's exact type and expose a public parameterless constructor.
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public sealed class ExcelConverterAttribute : Attribute
    {
        public Type ConverterType { get; }

        public ExcelConverterAttribute(Type converterType)
        {
            ArgumentNullException.ThrowIfNull(converterType);
            ConverterType = converterType;
        }
    }
}
