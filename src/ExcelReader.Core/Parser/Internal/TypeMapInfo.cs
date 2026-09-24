using System.Collections.Concurrent;

namespace ExcelReader.Core.Parser.Internal
{
    internal readonly struct TypeMapInfo<T>
        where T : allows ref struct
    {
        private readonly PropertyMap<T>[] _properties;
        private readonly Func<T>? _factory;
        private readonly bool _useDefault;
        private readonly ConcurrentDictionary<(StringComparer, HeaderNormalization), Dictionary<string, HeaderMatch<T>>> _lookupCache;
        private readonly ColumnBinding<T>[]? _indexBindings;

        internal TypeMapInfo(PropertyMap<T>[] properties, Func<T>? factory, bool useDefault)
        {
            _properties = properties;
            _factory = factory;
            _useDefault = useDefault;
            _lookupCache = new ConcurrentDictionary<(StringComparer, HeaderNormalization), Dictionary<string, HeaderMatch<T>>>();
        }

        internal TypeMapInfo(ColumnBinding<T>[] indexBindings, Func<T>? factory, bool useDefault)
        {
            _properties = [];
            _factory = factory;
            _useDefault = useDefault;
            _lookupCache = new ConcurrentDictionary<(StringComparer, HeaderNormalization), Dictionary<string, HeaderMatch<T>>>();
            _indexBindings = indexBindings;
        }

        internal int PropertyCount => _properties.Length;

        internal bool IsIndexBased => _indexBindings is not null;

        internal ColumnBinding<T>[] IndexBindings => _indexBindings!;

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
                throw new ExcelParseException(missing);
            }
        }

        internal static TypeMapInfo<T> MergeFluentOverAttributes(TypeMapInfo<T> fluent, TypeMapInfo<T> attributeFallback, StringComparer comparer, HeaderNormalization normalization)
        {
            if (fluent.IsIndexBased)
            {
                throw new InvalidOperationException(
                    "BuildWithAttributeFallback cannot merge a PropertyAt (index-based) map with attribute-driven properties: an index-based map has no header row to match attributes against. Use ExcelParser.Build instead.");
            }
            var configuredNames = new HashSet<string>(comparer);
            foreach (PropertyMap<T> property in fluent._properties)
            {
                foreach (string name in property.Names)
                {
                    configuredNames.Add(normalization.Apply(name));
                }
            }
            List<PropertyMap<T>> merged = [.. fluent._properties];
            foreach (PropertyMap<T> attributeProperty in attributeFallback._properties)
            {
                bool overridden = attributeProperty.Names.Any(name => configuredNames.Contains(normalization.Apply(name)));
                if (!overridden)
                {
                    merged.Add(attributeProperty);
                }
            }

            Func<T>? factory = fluent._useDefault ? attributeFallback._factory : fluent._factory;
            bool useDefault = fluent._useDefault && attributeFallback._useDefault;
            return new TypeMapInfo<T>([.. merged], factory, useDefault);
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
                    string name = normalization.Apply(property.Names[aliasIndex]);
                    if (!lookup.TryGetValue(name, out HeaderMatch<T> existing))
                    {
                        lookup.Add(name, new(propertyIndex, aliasIndex, property.Parser));
                    }
                    else if (existing.PropertyIndex != propertyIndex)
                    {
                        string[] existingNames = properties[existing.PropertyIndex].Names;
                        throw new InvalidOperationException(
                            $"Ambiguous header binding on {typeof(T)}: '{existingNames[existing.AliasIndex]}' (property bound as '{existingNames[0]}') and '{property.Names[aliasIndex]}' (property bound as '{property.Names[0]}') match the same header under the configured ColumnNameComparer and HeaderNormalization. Give each property distinct header names.");
                    }
                }
            }
            return lookup;
        }
    }
}
