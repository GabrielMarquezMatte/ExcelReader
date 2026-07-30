using System.Collections.Concurrent;

namespace ExcelReader.Core.Parser.Internal
{
    internal readonly struct TypeMapInfo<T>
#if NET9_0_OR_GREATER
        where T : allows ref struct
#endif
    {
        private readonly PropertyMap<T>[] _properties;
        // Null when _useDefault is true: a value type with no explicit parameterless constructor needs
        // no factory at all, since default(T) is exactly what `new T()` would have produced, and Build()
        // skips compiling one. Non-null (and always invoked instead of _useDefault) for every class T
        // and for a struct T that declares an explicit parameterless constructor (C# 10+), whose
        // user-written initializer logic default(T) would silently skip.
        private readonly Func<T>? _factory;
        private readonly bool _useDefault;
        private readonly ConcurrentDictionary<(StringComparer, HeaderNormalization), Dictionary<string, HeaderMatch<T>>> _lookupCache;

        internal TypeMapInfo(PropertyMap<T>[] properties, Func<T>? factory, bool useDefault)
        {
            _properties = properties;
            _factory = factory;
            _useDefault = useDefault;
            _lookupCache = new ConcurrentDictionary<(StringComparer, HeaderNormalization), Dictionary<string, HeaderMatch<T>>>();
        }

        internal int PropertyCount => _properties.Length;

        // Creates a fresh model instance per row without a `where T : new()` constraint, so types with
        // required members (which the new() constraint forbids) can still be parsed. For a plain struct
        // target (the common case for a zero-allocation row model), skips the compiled-factory delegate
        // call/struct-copy entirely — see Build()'s _useDefault computation.
        internal T CreateInstance()
        {
            return _useDefault ? default! : _factory!();
        }

        internal bool RequiresValue(int propertyIndex)
        {
            return _properties[propertyIndex].RequireValue;
        }

        internal string DisplayName(int propertyIndex)
        {
            return _properties[propertyIndex].Names[0];
        }

        // Throws if any [ExcelRequired] property was left unmatched after the header row was mapped.
        // unmatched[i] is int.MaxValue when property i found no header column (RowProjector's sentinel).
        internal void ValidateRequiredColumns(int[] unmatched)
        {
            List<string>? missing = null;
            for (int i = 0; i < _properties.Length; i++)
            {
                if (_properties[i].IsRequired && unmatched[i] == int.MaxValue)
                {
                    (missing ??= []).Add(_properties[i].Names[0]);
                }
            }
            if (missing is not null)
            {
                throw new InvalidOperationException(
                    $"Required column(s) not found in the header row: {string.Join(", ", missing)}.");
            }
        }

        internal bool TryFindHeader(string headerName, StringComparer comparer, HeaderNormalization normalization, out HeaderMatch<T> match)
        {
            PropertyMap<T>[] properties = _properties;
            var lookup = _lookupCache.GetOrAdd(
                (comparer, normalization),
                static (key, props) => BuildLookup(props, key.Item1, key.Item2),
                properties);
            return lookup.TryGetValue(headerName, out match);
        }

        private static Dictionary<string, HeaderMatch<T>> BuildLookup(PropertyMap<T>[] properties, StringComparer comparer, HeaderNormalization normalization)
        {
            Dictionary<string, HeaderMatch<T>> lookup = new(comparer);
            for (int propertyIndex = 0; propertyIndex < properties.Length; propertyIndex++)
            {
                PropertyMap<T> property = properties[propertyIndex];
                for (int aliasIndex = 0; aliasIndex < property.Names.Length; aliasIndex++)
                {
                    lookup.TryAdd(
                        normalization.Apply(property.Names[aliasIndex]),
                        new(propertyIndex, aliasIndex, property.Parser));
                }
            }
            return lookup;
        }
    }
}
