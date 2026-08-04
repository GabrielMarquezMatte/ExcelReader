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

        internal TypeMapInfo<T> Build()
        {
            return new TypeMapInfo<T>([.. _properties], _factory, useDefault: _factory is null);
        }
    }
}
