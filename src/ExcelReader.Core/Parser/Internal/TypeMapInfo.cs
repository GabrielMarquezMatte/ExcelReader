using System.Collections.Concurrent;

namespace ExcelReader.Core.Parser.Internal
{
    internal readonly struct TypeMapInfo<T> where T : new()
    {
        private readonly PropertyMap<T>[] _properties;
        private readonly ConcurrentDictionary<StringComparer, Dictionary<string, HeaderMatch<T>>> _lookupCache;

        internal TypeMapInfo(PropertyMap<T>[] properties)
        {
            _properties = properties;
            _lookupCache = new ConcurrentDictionary<StringComparer, Dictionary<string, HeaderMatch<T>>>();
        }

        internal int PropertyCount => _properties.Length;

        internal bool TryFindHeader(string headerName, StringComparer comparer, out HeaderMatch<T> match)
        {
            var lookup = _lookupCache.GetOrAdd(comparer, BuildLookup);
            return lookup.TryGetValue(headerName, out match);
        }

        private Dictionary<string, HeaderMatch<T>> BuildLookup(StringComparer comparer)
        {
            Dictionary<string, HeaderMatch<T>> lookup = new(comparer);
            for (int propertyIndex = 0; propertyIndex < _properties.Length; propertyIndex++)
            {
                PropertyMap<T> property = _properties[propertyIndex];
                for (int aliasIndex = 0; aliasIndex < property.Names.Length; aliasIndex++)
                {
                    lookup.TryAdd(
                        property.Names[aliasIndex],
                        new(propertyIndex, aliasIndex, property.Parser));
                }
            }
            return lookup;
        }
    }
}
