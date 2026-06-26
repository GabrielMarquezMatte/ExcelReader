using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.ExceptionServices;

namespace ExcelReader.Core.Parser.Internal
{
    internal static class TypeMapper<T>
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
                ExcelRequiredAttribute? requiredAttr = prop.GetCustomAttribute<ExcelRequiredAttribute>();
                bool isRequired = requiredAttr is not null;
                bool requireValue = isRequired && !requiredAttr!.AllowEmpty;
                ExcelConverterAttribute? converterAttr = prop.GetCustomAttribute<ExcelConverterAttribute>();
                ColumnParser<T>? parser = converterAttr is not null
                    ? ColumnParserFactory.BuildConverter<T>(prop, converterAttr.ConverterType)
                    : ColumnParserFactory.Build<T>(prop);
                if (parser is null)
                {
                    if (isRequired)
                    {
                        // A required column with no parser could never bind, so its requirement would
                        // be impossible to satisfy — surface that as a configuration error up front.
                        throw new InvalidOperationException(
                            $"Property '{typeof(T).Name}.{prop.Name}' is marked [ExcelRequired] but its type '{prop.PropertyType}' has no parser. Add an [ExcelConverter] for it.");
                    }
                    continue;
                }
                ExcelColumnAttribute[] attrs = [.. prop.GetCustomAttributes<ExcelColumnAttribute>()];
                string[] names = attrs.Length == 0
                    ? [prop.Name]
                    : [.. attrs.Select(static attr => attr.Name)];
                propertyMaps.Add(new PropertyMap<T>(names, parser, isRequired, requireValue));
            }

            // Compiled once per type — the per-row instance factory that replaces `new T()`, so the
            // parser no longer needs a `where T : new()` constraint (and types with required members work).
            Func<T> factory = Expression.Lambda<Func<T>>(Expression.New(typeof(T))).Compile();
            return new TypeMapInfo<T>([.. propertyMaps], factory);
        }
    }
}
