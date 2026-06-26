using System.Collections.Concurrent;

namespace ExcelReader.Core.Parser.Internal
{
    internal readonly struct TypeMapInfo<T> where T : new()
    {
        private readonly PropertyMap<T>[] _properties;
        private readonly ConcurrentDictionary<(StringComparer, HeaderNormalization), Dictionary<string, HeaderMatch<T>>> _lookupCache;

        internal TypeMapInfo(PropertyMap<T>[] properties)
        {
            _properties = properties;
            _lookupCache = new ConcurrentDictionary<(StringComparer, HeaderNormalization), Dictionary<string, HeaderMatch<T>>>();
        }

        internal int PropertyCount => _properties.Length;

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
