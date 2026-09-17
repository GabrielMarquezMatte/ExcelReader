using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;

namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// How CSV columns bind to the properties of <typeparamref name="TModel"/>, for
    /// <c>Excel.AggregateCsvParallelAsync</c>. Create one with <see cref="CsvModelMap"/>.
    /// </summary>
    /// <typeparam name="TModel">The model type. May be a <see langword="ref struct"/> with <see cref="ReadOnlySpan{T}"/> of <see cref="byte"/> properties.</typeparam>
    public sealed class CsvModelMap<TModel>
        where TModel : allows ref struct
    {
        internal CsvModelMap(TypeMapInfo<TModel> info, ExcelParserConfig config)
        {
            Info = info;
            Config = config;
        }

        internal TypeMapInfo<TModel> Info { get; }

        internal ExcelParserConfig Config { get; }
    }

    /// <summary>Creates <see cref="CsvModelMap{TModel}"/> instances.</summary>
    /// <remarks>
    /// <see cref="ExcelParserConfig.HeaderRow"/> is ignored; the header position comes from
    /// <see cref="CsvParallelOptions.HeaderRow"/>. The other settings control header matching, culture
    /// and parse failures.
    /// </remarks>
    public static class CsvModelMap
    {
        /// <summary>Uses the map the <c>[ExcelSerializable]</c> source generator emitted for <typeparamref name="T"/>. Trimming and AOT safe.</summary>
        /// <typeparam name="T">The model type.</typeparam>
        /// <param name="config">Header matching, culture and parse-failure options. Defaults to a new <see cref="ExcelParserConfig"/>.</param>
        /// <returns>The map.</returns>
        public static CsvModelMap<T> Generated<T>(ExcelParserConfig? config = null)
            where T : IExcelRowMap<T>, allows ref struct
        {
            var builder = new ExcelRowMapBuilder<T>();
            T.ConfigureExcelRowMap(builder);
            return new CsvModelMap<T>(builder.Build(), config ?? new ExcelParserConfig());
        }

        /// <summary>Builds a map with <see cref="ExcelRowMapBuilder{T}"/>. Trimming and AOT safe.</summary>
        /// <typeparam name="T">The model type.</typeparam>
        /// <param name="configure">Configures the builder.</param>
        /// <param name="config">Header matching, culture and parse-failure options. Defaults to a new <see cref="ExcelParserConfig"/>.</param>
        /// <returns>The map.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="configure"/> is <see langword="null"/>.</exception>
        public static CsvModelMap<T> Build<T>(Action<ExcelRowMapBuilder<T>> configure, ExcelParserConfig? config = null)
            where T : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(configure);
            var builder = new ExcelRowMapBuilder<T>();
            configure(builder);
            return new CsvModelMap<T>(builder.Build(), config ?? new ExcelParserConfig());
        }

        /// <summary>Builds a map by reflecting over <typeparamref name="T"/>'s <c>[ExcelColumn]</c>/<c>[ExcelRequired]</c>/<c>[ExcelConverter]</c> attributes.</summary>
        /// <typeparam name="T">The model type.</typeparam>
        /// <param name="config">Header matching, culture and parse-failure options. Defaults to a new <see cref="ExcelParserConfig"/>.</param>
        /// <returns>The map.</returns>
        [RequiresUnreferencedCode("FromAttributes reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("FromAttributes binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        public static CsvModelMap<T> FromAttributes<T>(ExcelParserConfig? config = null)
            where T : allows ref struct
        {
            return new CsvModelMap<T>(TypeMapper<T>.GetCsvInfo(), config ?? new ExcelParserConfig());
        }
    }
}
