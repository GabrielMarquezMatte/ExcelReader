using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace ExcelReader.Core.Parser.Internal
{
    // Every public entry point that reaches this type (ExcelParser<T>, RefParser.ParseNamed,
    // ExcelFluentParser<T>.WithAttributeFallback) already carries its own matching
    // [RequiresUnreferencedCode]/[RequiresDynamicCode] pair - this class-level pair is what makes that
    // annotation actually cover the reflection this class performs internally, instead of just
    // documenting an assumption at the outer boundary.
    [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
    [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
    internal static class TypeMapper<T>
#if NET9_0_OR_GREATER
        where T : allows ref struct
#endif
    {
        // ExecutionAndPublication already caches a thrown build exception and re-throws it (original
        // stack trace preserved) on every subsequent .Value access — no need to do that by hand.
        private static readonly Lazy<TypeMapInfo<T>> _info =
            new(static () => Build(csvTextDates: false), LazyThreadSafetyMode.ExecutionAndPublication);

        // Separate cache for the CSV parser: identical to _info except DateTime/DateOnly parse text
        // instead of an Excel serial number. Built lazily so non-CSV callers never pay for it.
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

        // csvTextDates only changes how DateTime/DateOnly properties parse — for a T with none, the two
        // maps would be built byte-for-byte identical, so this reuses _info's already-reflected,
        // already-compiled map instead of re-reflecting every property and recompiling every setter.
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
                if (prop.GetSetMethod() is null)
                {
                    continue;
                }
                if (Attribute.IsDefined(prop, typeof(ExcelIgnoreAttribute)))
                {
                    continue;
                }
                ExcelRequiredAttribute? requiredAttr = prop.GetCustomAttribute<ExcelRequiredAttribute>();
                bool isRequired = requiredAttr is not null;
                bool requireValue = isRequired && !requiredAttr!.AllowEmpty;
                ExcelConverterAttribute? converterAttr = prop.GetCustomAttribute<ExcelConverterAttribute>();
                ColumnParser<T>? parser = converterAttr is not null
                    ? ColumnParserFactory.BuildConverter<T>(prop, converterAttr.ConverterType)
                    : ColumnParserFactory.Build<T>(prop, csvTextDates);
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

            // A value type with no explicit parameterless constructor needs no factory at all: default(T)
            // is byte-for-byte what `new T()` would produce, so CreateInstance() can skip the delegate
            // call (and the struct copy `Func<T>` returns by value) entirely. Reflection only reports a
            // struct's parameterless constructor when the type explicitly declares one (C# 10+); the
            // ordinary implicit one is invisible here, so this stays false for the common plain-struct
            // case and true (needing the compiled factory) only for a genuine user-defined constructor.
            bool useDefault = typeof(T).IsValueType && typeof(T).GetConstructor(Type.EmptyTypes) is null;
            // Activator.CreateInstance<T>() binds the parameterless constructor directly — no
            // Expression.Compile() dynamic-method emission. Requires a public parameterless
            // constructor, matching the GetConstructor(Type.EmptyTypes) lookup above (public-only).
            Func<T>? factory = useDefault ? null : static () => Activator.CreateInstance<T>();
            return new TypeMapInfo<T>([.. propertyMaps], factory, useDefault);
        }
    }
}
