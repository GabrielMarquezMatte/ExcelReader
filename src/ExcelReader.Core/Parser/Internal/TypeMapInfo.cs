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
        // Non-null exactly for a fluent map built entirely from ExcelRowMapBuilder<T>.PropertyAt: fixed
        // column index, no header row involved at all. Already sorted by Column — the shape RowProjector/
        // CsvRowProjector need directly, with no per-row header lookup step (see IsIndexBased).
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

        // True for a map built purely by ExcelRowMapBuilder<T>.PropertyAt (§4.4.2): no header row exists
        // to wait for, so RowProjector/CsvRowProjector build the column map immediately instead of at
        // ProjectionRules.ClassifyRow's usual header-row step (R-C3 — HeaderRow is never repurposed to
        // mean this).
        internal bool IsIndexBased => _indexBindings is not null;

        internal ColumnBinding<T>[] IndexBindings => _indexBindings!;

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
        // A header missing a required column is a defect in the file, not a caller mistake, so this
        // throws ExcelParseException — the same type used for a per-row parse/required-value failure —
        // rather than InvalidOperationException, which would misattribute the fault to the caller.
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

        // Fluent overrides attribute, per property (§4.4.3): a property configured in `fluent` fully
        // replaces whatever attribute-driven property shares one of its header names, regardless of
        // comparer/normalization used later at parse time — the same identity a header row itself would
        // use to pick between two same-named bindings. A property the builder never touched keeps its
        // attribute-driven behavior untouched.
        internal static TypeMapInfo<T> MergeFluentOverAttributes(TypeMapInfo<T> fluent, TypeMapInfo<T> attributeFallback, StringComparer comparer, HeaderNormalization normalization)
        {
            // An index-based map (ExcelRowMapBuilder<T>.PropertyAt) has no header row to match
            // attribute-driven properties against — it stores its bindings in _indexBindings and leaves
            // _properties empty, so silently ignoring them here would drop every PropertyAt binding
            // rather than merge it. The two shapes can't compose; fail loud instead.
            if (fluent.IsIndexBased)
            {
                throw new InvalidOperationException(
                    "WithAttributeFallback cannot merge a PropertyAt (index-based) map with attribute-driven properties: an index-based map has no header row to match attributes against. Use the ExcelFluentParser<T> constructor instead.");
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

            // A fluent builder that never called Factory() has _useDefault = true (see
            // ExcelRowMapBuilder<T>.Factory's doc comment) — fall back to the attribute map's factory
            // rather than silently defaulting a class T to null.
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
                    lookup.TryAdd(
                        normalization.Apply(property.Names[aliasIndex]),
                        new(propertyIndex, aliasIndex, property.Parser));
                }
            }
            return lookup;
        }
    }
}
