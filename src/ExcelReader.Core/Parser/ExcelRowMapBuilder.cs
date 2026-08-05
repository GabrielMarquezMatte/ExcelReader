using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Accumulates a row-to-property binding for <typeparamref name="T"/>, fed either by the source
    /// generator (feature A) or by a hand-written <see cref="IExcelRowMap{T}"/> implementation. Mirrors,
    /// without reflection, exactly what <c>TypeMapper&lt;T&gt;.Build()</c> produces from attributes.
    /// </summary>
    /// <typeparam name="T">The row model type being mapped.</typeparam>
    public sealed class ExcelRowMapBuilder<T>
#if NET9_0_OR_GREATER
        where T : allows ref struct
#endif
    {
        private readonly List<PropertyMap<T>> _properties = [];
        private readonly List<ColumnBinding<T>> _indexBindings = [];
        private Func<T>? _factory;

        /// <summary>
        /// Binds a property to one or more header names, using <paramref name="read"/> to convert the
        /// matched cell and <paramref name="set"/> to assign the result.
        /// </summary>
        /// <typeparam name="TValue">The property's value type.</typeparam>
        /// <param name="names">The header name(s) that bind to this property; the first is used in error messages.</param>
        /// <param name="read">Converts a matched, non-empty cell into a <typeparamref name="TValue"/>; see <see cref="ExcelCellReaders"/> for the built-in readers.</param>
        /// <param name="set">Assigns the read value to the property.</param>
        /// <param name="isRequired">Whether the header must be present.</param>
        /// <param name="requireValue">Whether every data row's cell for this column must also be non-empty.</param>
        /// <returns>This builder, for chaining.</returns>
        public ExcelRowMapBuilder<T> Property<TValue>(
            string[] names,
            ExcelCellReader<TValue> read,
            ExcelPropertySetter<T, TValue> set,
            bool isRequired = false,
            bool requireValue = false)
        {
            ArgumentNullException.ThrowIfNull(names);
            ArgumentNullException.ThrowIfNull(read);
            ArgumentNullException.ThrowIfNull(set);
            bool parser(ref T model, in Cell cell, bool isDate1904, IFormatProvider provider)
            {
                if (!read(in cell, isDate1904, provider, out TValue value))
                {
                    return false;
                }
                set(ref model, value);
                return true;
            }
            _properties.Add(new PropertyMap<T>(names, parser, isRequired, requireValue));
            return this;
        }

        /// <summary>
        /// Binds a <see cref="Nullable{TValue}"/> property to one or more header names, using
        /// <paramref name="read"/> — one of the non-nullable built-in readers, e.g.
        /// <see cref="ExcelCellReaders.Bool"/> or <see cref="ExcelCellReaders.Parsable{TValue}"/> — to
        /// convert the matched cell, then wraps the result in <see cref="Nullable{TValue}"/> before
        /// calling <paramref name="set"/>. There is no nullable-returning counterpart of the built-in
        /// readers to plug into <see cref="Property{TValue}"/> directly, because the wrapping step is
        /// identical for every value type — this method does it once, generically.
        /// </summary>
        /// <typeparam name="TValue">The property's underlying (non-nullable) value type.</typeparam>
        /// <param name="names">The header name(s) that bind to this property; the first is used in error messages.</param>
        /// <param name="read">Converts a matched, non-empty cell into a <typeparamref name="TValue"/>.</param>
        /// <param name="set">Assigns the read value, wrapped in <see cref="Nullable{TValue}"/>, to the property.</param>
        /// <param name="isRequired">Whether the header must be present.</param>
        /// <param name="requireValue">Whether every data row's cell for this column must also be non-empty.</param>
        /// <returns>This builder, for chaining.</returns>
        public ExcelRowMapBuilder<T> PropertyNullable<TValue>(
            string[] names,
            ExcelCellReader<TValue> read,
            ExcelPropertySetter<T, TValue?> set,
            bool isRequired = false,
            bool requireValue = false)
            where TValue : struct
        {
            ArgumentNullException.ThrowIfNull(names);
            ArgumentNullException.ThrowIfNull(read);
            ArgumentNullException.ThrowIfNull(set);
            bool parser(ref T model, in Cell cell, bool isDate1904, IFormatProvider provider)
            {
                if (!read(in cell, isDate1904, provider, out TValue value))
                {
                    return false;
                }
                set(ref model, value);
                return true;
            }
            _properties.Add(new PropertyMap<T>(names, parser, isRequired, requireValue));
            return this;
        }

        /// <summary>
        /// Binds a header to a single already-fused read+assign delegate, instead of the separate
        /// <see cref="ExcelCellReader{TValue}"/> + <see cref="ExcelPropertySetter{TModel, TValue}"/> pair
        /// <see cref="Property{TValue}"/>/<see cref="PropertyNullable{TValue}"/> compose internally. Every
        /// bound cell costs one indirect call instead of two (read, then set) — the source generator uses
        /// this for <c>[ExcelSerializable]</c> models. Prefer <see cref="Property{TValue}"/> for hand-written
        /// maps unless the extra call genuinely matters for your workload; it's a small, row-count-proportional
        /// cost, not a correctness concern either way.
        /// </summary>
        /// <param name="names">The header name(s) that bind to this property; the first is used in error messages.</param>
        /// <param name="parse">Reads a matched, non-empty cell and assigns it directly onto the model.</param>
        /// <param name="isRequired">Whether the header must be present.</param>
        /// <param name="requireValue">Whether every data row's cell for this column must also be non-empty.</param>
        /// <returns>This builder, for chaining.</returns>
        public ExcelRowMapBuilder<T> PropertyRaw(
            string[] names,
            ExcelRowParser<T> parse,
            bool isRequired = false,
            bool requireValue = false)
        {
            ArgumentNullException.ThrowIfNull(names);
            ArgumentNullException.ThrowIfNull(parse);
            // ExcelRowParser<T> and the internal ColumnParser<T> share an identical invoke signature, so
            // this constructs a new delegate instance targeting the same method/closure directly — not a
            // wrapper that calls through parse.Invoke(...). No extra indirection at row-parse time.
            _properties.Add(new PropertyMap<T>(names, new ColumnParser<T>(parse), isRequired, requireValue));
            return this;
        }

        /// <summary>
        /// Binds a property to one or more header names via a <c>[ExcelConverter]</c>-style converter,
        /// for value types none of the built-in <see cref="ExcelCellReaders"/> cover. The generator emits
        /// <c>new MyConverter()</c> for <paramref name="converter"/> — nothing here uses <see cref="Activator"/>.
        /// </summary>
        /// <typeparam name="TValue">The property's value type.</typeparam>
        /// <param name="names">The header name(s) that bind to this property; the first is used in error messages.</param>
        /// <param name="converter">Converts a matched, non-empty cell into a <typeparamref name="TValue"/>.</param>
        /// <param name="set">Assigns the converted value to the property.</param>
        /// <param name="isRequired">Whether the header must be present.</param>
        /// <param name="requireValue">Whether every data row's cell for this column must also be non-empty.</param>
        /// <returns>This builder, for chaining.</returns>
        public ExcelRowMapBuilder<T> Converted<TValue>(
            string[] names,
            IExcelCellConverter<TValue> converter,
            ExcelPropertySetter<T, TValue> set,
            bool isRequired = false,
            bool requireValue = false)
        {
            ArgumentNullException.ThrowIfNull(names);
            ArgumentNullException.ThrowIfNull(converter);
            ArgumentNullException.ThrowIfNull(set);
            bool parser(ref T model, in Cell cell, bool isDate1904, IFormatProvider provider)
            {
                if (!converter.TryConvert(in cell, isDate1904, provider, out TValue value))
                {
                    return false;
                }
                set(ref model, value);
                return true;
            }
            _properties.Add(new PropertyMap<T>(names, parser, isRequired, requireValue));
            return this;
        }

        /// <summary>
        /// Binds a property to a fixed 0-based column index instead of a header name, so rows with no
        /// header at all (or a header the caller doesn't want matched by text) can still be mapped.
        /// </summary>
        /// <remarks>
        /// A builder that uses this method exclusively produces a map with no header-row step at all —
        /// the first row is already a data row. Mixing this with <see cref="Property{TValue}"/>/
        /// <see cref="PropertyNullable{TValue}"/>/<see cref="Converted{TValue}"/> on the same builder is
        /// not supported, since one map can't both wait for a header row and skip it.
        /// </remarks>
        /// <typeparam name="TValue">The property's value type.</typeparam>
        /// <param name="columnIndex">The 0-based column index to read from.</param>
        /// <param name="read">Converts a matched, non-empty cell into a <typeparamref name="TValue"/>; see <see cref="ExcelCellReaders"/> for the built-in readers.</param>
        /// <param name="set">Assigns the read value to the property.</param>
        /// <param name="requireValue">Whether the cell at <paramref name="columnIndex"/> must be non-empty in every data row.</param>
        /// <returns>This builder, for chaining.</returns>
        public ExcelRowMapBuilder<T> PropertyAt<TValue>(
            int columnIndex,
            ExcelCellReader<TValue> read,
            ExcelPropertySetter<T, TValue> set,
            bool requireValue = false)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
            ArgumentNullException.ThrowIfNull(read);
            ArgumentNullException.ThrowIfNull(set);
            bool parser(ref T model, in Cell cell, bool isDate1904, IFormatProvider provider)
            {
                if (!read(in cell, isDate1904, provider, out TValue value))
                {
                    return false;
                }
                set(ref model, value);
                return true;
            }
            _indexBindings.Add(new ColumnBinding<T>(columnIndex, parser, requireValue, $"Column {columnIndex}"));
            return this;
        }

        /// <summary>
        /// Sets the factory used to create a fresh <typeparamref name="T"/> instance per row.
        /// <see langword="null"/> means <c>default(T)</c> — use this for a plain struct with no
        /// explicit parameterless constructor, matching <c>TypeMapper&lt;T&gt;.Build()</c>'s own rule.
        /// Not calling this method at all has the same effect as calling it with <see langword="null"/>,
        /// which is wrong for any class or struct with an explicit constructor to run — always call it
        /// for those.
        /// </summary>
        /// <param name="factory">The instance factory, or <see langword="null"/> to use <c>default(T)</c>.</param>
        /// <returns>This builder, for chaining.</returns>
        public ExcelRowMapBuilder<T> Factory(Func<T>? factory)
        {
            _factory = factory;
            return this;
        }

        // requireFactory: false for ExcelFluentParser<T>.WithAttributeFallback's fluent-only builder —
        // that path builds this map before merging it with the attribute-driven one, and
        // TypeMapInfo<T>.MergeFluentOverAttributes already falls back to the attribute map's factory
        // when this one has none (see its own comment on _useDefault). Checking here unconditionally
        // would reject that legitimate, already-rescued case before the merge ever runs.
        internal TypeMapInfo<T> Build(bool requireFactory = true)
        {
            // A null factory means default(T); for a reference type that's always null, and every row
            // would throw NullReferenceException on the first property assignment instead of at Build()
            // time, where the real mistake (forgetting to call Factory(...)) is far easier to see.
            if (requireFactory && _factory is null && !typeof(T).IsValueType)
            {
                throw new InvalidOperationException(
                    $"ExcelRowMapBuilder<{typeof(T)}> has no factory: {typeof(T)} is a reference type, so default(T) is null and every row would fail on the first property assignment. Call Factory(...) with a real instance factory.");
            }
            if (_indexBindings.Count == 0)
            {
                return new TypeMapInfo<T>([.. _properties], _factory, useDefault: _factory is null);
            }
            if (_properties.Count > 0)
            {
                throw new InvalidOperationException(
                    "Cannot mix Property/PropertyNullable/Converted with PropertyAt on the same ExcelRowMapBuilder<T>.");
            }
            ColumnBinding<T>[] bindings = [.. _indexBindings];
            Array.Sort(bindings, static (left, right) => left.Column.CompareTo(right.Column));
            return new TypeMapInfo<T>(bindings, _factory, useDefault: _factory is null);
        }
    }
}
