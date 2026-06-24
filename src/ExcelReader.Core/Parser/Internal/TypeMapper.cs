using System.Collections.Concurrent;
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

    internal readonly struct PropertyMap<T> where T : new()
    {
        internal PropertyMap(string[] names, ColumnParser<T> parser)
        {
            Names = names;
            Parser = parser;
        }

        internal string[] Names { get; }
        internal ColumnParser<T> Parser { get; }
    }

    internal readonly struct HeaderMatch<T> where T : new()
    {
        internal HeaderMatch(int propertyIndex, int aliasIndex, ColumnParser<T> parser)
        {
            PropertyIndex = propertyIndex;
            AliasIndex = aliasIndex;
            Parser = parser;
        }

        internal int PropertyIndex { get; }
        internal int AliasIndex { get; }
        internal ColumnParser<T> Parser { get; }
    }
}
