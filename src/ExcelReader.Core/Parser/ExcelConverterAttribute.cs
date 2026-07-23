namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Binds a custom <see cref="IExcelCellConverter{T}"/> to a property. The converter type must
    /// implement <see cref="IExcelCellConverter{T}"/> for the property's exact type and expose a public
    /// parameterless constructor.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public sealed class ExcelConverterAttribute : Attribute
    {
        /// <summary>The <see cref="IExcelCellConverter{T}"/> implementation to use for this property.</summary>
        public Type ConverterType { get; }

        /// <summary>Creates an attribute that binds the given converter type to the property.</summary>
        public ExcelConverterAttribute(Type converterType)
        {
            ArgumentNullException.ThrowIfNull(converterType);
            ConverterType = converterType;
        }
    }
}
