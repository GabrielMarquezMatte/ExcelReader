using System.Reflection;
using System.Runtime.ExceptionServices;

namespace ExcelReader.Core.Parser.Internal
{
    internal static class TypeMapper<T> where T : new()
    {
        private static readonly Lazy<TypeMapInfo<T>> _info =
            new(BuildSafe, LazyThreadSafetyMode.ExecutionAndPublication);

        internal static TypeMapInfo<T> GetInfo()
        {
            return _info.Value;
        }

        private static TypeMapInfo<T> BuildSafe()
        {
            try
            {
                return Build();
            }
            catch (Exception ex)
            {
                // Capture here so the exception is re-thrown with original stack trace on every call.
                ExceptionDispatchInfo.Capture(ex).Throw();
                throw; // unreachable; satisfies compiler
            }
        }

        private static TypeMapInfo<T> Build()
        {
            PropertyInfo[] properties = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance);

            var propertyMaps = new List<PropertyMap<T>>(properties.Length);

            foreach (PropertyInfo prop in properties)
            {
                if (!prop.CanWrite || prop.GetSetMethod() is null)
                {
                    continue;
                }
                var parser = ColumnParserFactory.Build<T>(prop);
                if (parser is null)
                {
                    continue;
                }
                ExcelColumnAttribute[] attrs = [.. prop.GetCustomAttributes<ExcelColumnAttribute>()];
                string[] names = attrs.Length == 0
                    ? [prop.Name]
                    : [.. attrs.Select(static attr => attr.Name)];
                propertyMaps.Add(new PropertyMap<T>(names, parser));
            }

            return new TypeMapInfo<T>([.. propertyMaps]);
        }
    }
}
