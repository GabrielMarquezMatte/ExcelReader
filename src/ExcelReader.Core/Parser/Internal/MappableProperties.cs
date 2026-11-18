using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelReader.Core.Parser.Internal
{
    internal static class MappableProperties
    {
        // Most-derived type first, so a property hidden with `new` loses to the one hiding it, as in the source generator.
        [RequiresUnreferencedCode("Reflects over the type's public properties, which trimming may remove.")]
        internal static List<PropertyInfo> Of(Type type)
        {
            var seenNames = new HashSet<string>(StringComparer.Ordinal);
            var properties = new List<PropertyInfo>();
            for (Type? current = type; current is not null; current = current.BaseType)
            {
                foreach (PropertyInfo prop in current.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
                {
                    if (prop.GetIndexParameters().Length == 0 && !Attribute.IsDefined(prop, typeof(ExcelIgnoreAttribute)) && seenNames.Add(prop.Name))
                    {
                        properties.Add(prop);
                    }
                }
            }
            return properties;
        }
    }
}
