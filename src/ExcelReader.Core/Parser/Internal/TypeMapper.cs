using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelReader.Core.Parser.Internal
{
    [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
    [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
    internal static class TypeMapper<T>
        where T : allows ref struct
    {
        private static readonly Lazy<TypeMapInfo<T>> _info =
            new(static () => Build(csvTextDates: false), LazyThreadSafetyMode.ExecutionAndPublication);

        private static readonly Lazy<TypeMapInfo<T>> _csvInfo =
            new(BuildCsvInfo, LazyThreadSafetyMode.ExecutionAndPublication);

        internal static TypeMapInfo<T> GetInfo()
        {
            return _info.Value;
        }

        internal static TypeMapInfo<T> GetCsvInfo()
        {
            return _csvInfo.Value;
        }

        private static TypeMapInfo<T> BuildCsvInfo()
        {
            return HasDateProperty() ? Build(csvTextDates: true) : _info.Value;
        }

        private static bool HasDateProperty()
        {
            return typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance).Any(static prop =>
            {
                Type effective = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                return effective == typeof(DateTime) || effective == typeof(DateOnly);
            });
        }

        private static TypeMapInfo<T> Build(bool csvTextDates)
        {
            PropertyInfo[] properties = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance);

            var propertyMaps = new List<PropertyMap<T>>(properties.Length);

            foreach (PropertyInfo prop in properties)
            {
                if (Attribute.IsDefined(prop, typeof(ExcelIgnoreAttribute)))
                {
                    continue;
                }
                ExcelRequiredAttribute? requiredAttr = prop.GetCustomAttribute<ExcelRequiredAttribute>();
                bool isRequired = requiredAttr is not null;
                if (prop.GetSetMethod() is null)
                {
                    if (isRequired)
                    {
                        throw new InvalidOperationException(
                            $"Property '{typeof(T).Name}.{prop.Name}' is marked [ExcelRequired] but has no public set or init accessor, so it can never be read. Add one, or remove [ExcelRequired].");
                    }
                    continue;
                }
                bool requireValue = isRequired && !requiredAttr!.AllowEmpty;
                ExcelConverterAttribute? converterAttr = prop.GetCustomAttribute<ExcelConverterAttribute>();
                ColumnParser<T>? parser = converterAttr is not null
                    ? ColumnParserFactory.BuildConverter<T>(prop, converterAttr.ConverterType)
                    : ColumnParserFactory.Build<T>(prop, csvTextDates);
                if (parser is null)
                {
                    if (isRequired)
                    {
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

            bool useDefault = typeof(T).IsValueType && typeof(T).GetConstructor(Type.EmptyTypes) is null;
            Func<T>? factory = useDefault ? null : static () => Activator.CreateInstance<T>();
            return new TypeMapInfo<T>([.. propertyMaps], factory, useDefault);
        }
    }
}
