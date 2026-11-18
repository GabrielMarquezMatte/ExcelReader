using System.Reflection;

namespace ExcelReader.NetFramework.Parser.Internal
{
    internal static class MappedProperties
    {
        internal static IEnumerable<PropertyInfo> Of(Type type)
        {
            return type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.GetIndexParameters().Length == 0)
                .GroupBy(p => p.Name, StringComparer.Ordinal)
                .Select(g => g.Aggregate((a, b) => b.DeclaringType!.IsSubclassOf(a.DeclaringType!) ? b : a))
                .Where(p => !Attribute.IsDefined(p, typeof(ExcelIgnoreAttribute)));
        }
    }
}
