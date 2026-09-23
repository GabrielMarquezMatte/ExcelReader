using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Parses rows from an Excel or CSV reader into instances of <typeparamref name="T"/>. Create one with
    /// <see cref="ExcelParser"/>: <see cref="ExcelParser.FromAttributes{T}"/> reflects over
    /// <c>[ExcelColumn]</c>/<c>[ExcelRequired]</c>/<c>[ExcelConverter]</c> attributes,
    /// <see cref="ExcelParser.Generated{T}"/> uses the <c>[ExcelSerializable]</c> source-generated map, and
    /// <see cref="ExcelParser.Build{T}"/> takes a map configured at runtime.
    /// </summary>
    /// <typeparam name="T">
    /// The model type to bind each row to. May be a <see langword="ref struct"/>, including one with a
    /// <c>ReadOnlySpan&lt;byte&gt;</c> property. A <c>struct T</c> consumed via a direct <c>foreach</c>
    /// skips the per-row model allocation a class <typeparamref name="T"/> requires.
    /// </typeparam>
    public sealed class ExcelParser<T>
        where T : allows ref struct
    {
        private readonly ExcelParserConfig _config;
        private readonly Func<TypeMapInfo<T>> _info;
        private readonly Func<TypeMapInfo<T>> _csvInfo;

        internal ExcelParser(ExcelParserConfig config, Func<TypeMapInfo<T>> info, Func<TypeMapInfo<T>> csvInfo)
        {
            _config = config;
            _info = info;
            _csvInfo = csvInfo;
        }

        /// <summary>Parses the rows of an XLSX reader into <typeparamref name="T"/> instances, lazily as the result is enumerated.</summary>
        /// <param name="reader">The XLSX reader to pull rows from.</param>
        /// <returns>An enumerable that yields one <typeparamref name="T"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T> Parse(XlsxReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T>(reader, _config, _info());
        }

        /// <summary>Parses the rows of an XLS reader into <typeparamref name="T"/> instances, lazily as the result is enumerated.</summary>
        /// <param name="reader">The XLS reader to pull rows from.</param>
        /// <returns>An enumerable that yields one <typeparamref name="T"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T, XlsReader, XlsReader.Enumerator> Parse(XlsReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, XlsReader, XlsReader.Enumerator>(reader, _config, _info());
        }

        /// <summary>Parses the rows of an XLSB reader into <typeparamref name="T"/> instances, lazily as the result is enumerated.</summary>
        /// <param name="reader">The XLSB reader to pull rows from.</param>
        /// <returns>An enumerable that yields one <typeparamref name="T"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T, XlsbReader, XlsbReader.Enumerator> Parse(XlsbReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, XlsbReader, XlsbReader.Enumerator>(reader, _config, _info());
        }

        /// <summary>Parses the rows of a format-agnostic reader (e.g. one returned by <c>Excel.Open</c>) into <typeparamref name="T"/> instances, lazily as the result is enumerated. Lets callers avoid pattern-matching the concrete reader type; dispatches through the interface enumerator.</summary>
        /// <param name="reader">The reader to pull rows from.</param>
        /// <returns>An enumerable that yields one <typeparamref name="T"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator> Parse(IExcelRowReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator>(reader, _config, _info());
        }

        /// <summary>Parses the rows of a CSV reader into <typeparamref name="T"/> instances, lazily as the result is enumerated.</summary>
        /// <param name="reader">The CSV reader to pull rows from.</param>
        /// <returns>An enumerable that yields one <typeparamref name="T"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        /// <remarks>
        /// Uses a specialized enumerable (dense field binding, single-pass projection) rather than the
        /// generic <see cref="ExcelEnumerable{T}"/>, since CSV rows have no gaps or styles. Prefer this
        /// concrete overload over <see cref="Parse(IExcelRowReader)"/> for CSV — holding the reader as
        /// <see cref="IExcelRowReader"/> instead routes through the generic path.
        /// </remarks>
        public CsvEnumerable<T> Parse(CsvReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new CsvEnumerable<T>(reader, _config, _csvInfo());
        }

        /// <summary>Parses the rows of an XLSX reader into <typeparamref name="T"/> instances for asynchronous enumeration.</summary>
        /// <param name="reader">The XLSX reader to pull rows from.</param>
        /// <param name="ct">A token to cancel the enumeration.</param>
        /// <returns>An enumerable that lazily parses and yields one <typeparamref name="T"/> per data row as it is asynchronously enumerated.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T> ParseAsync(XlsxReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T>(reader, _config, _info(), ct);
        }

        /// <summary>Parses the rows of an XLS reader into <typeparamref name="T"/> instances for asynchronous enumeration.</summary>
        /// <param name="reader">The XLS reader to pull rows from.</param>
        /// <param name="ct">A token to cancel the enumeration.</param>
        /// <returns>An enumerable that lazily parses and yields one <typeparamref name="T"/> per data row as it is asynchronously enumerated.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T, XlsReader, XlsReader.Enumerator> ParseAsync(XlsReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, XlsReader, XlsReader.Enumerator>(reader, _config, _info(), ct);
        }

        /// <summary>Parses the rows of an XLSB reader into <typeparamref name="T"/> instances for asynchronous enumeration.</summary>
        /// <param name="reader">The XLSB reader to pull rows from.</param>
        /// <param name="ct">A token to cancel the enumeration.</param>
        /// <returns>An enumerable that lazily parses and yields one <typeparamref name="T"/> per data row as it is asynchronously enumerated.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T, XlsbReader, XlsbReader.Enumerator> ParseAsync(XlsbReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, XlsbReader, XlsbReader.Enumerator>(reader, _config, _info(), ct);
        }

        /// <summary>Parses the rows of a format-agnostic reader (e.g. one returned by <c>Excel.Open</c>) into <typeparamref name="T"/> instances for asynchronous enumeration.</summary>
        /// <param name="reader">The reader to pull rows from.</param>
        /// <param name="ct">A token to cancel the enumeration.</param>
        /// <returns>An enumerable that lazily parses and yields one <typeparamref name="T"/> per data row as it is asynchronously enumerated.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator> ParseAsync(IExcelRowReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator>(reader, _config, _info(), ct);
        }

        /// <summary>Parses the rows of a CSV reader into <typeparamref name="T"/> instances for asynchronous enumeration.</summary>
        /// <param name="reader">The CSV reader to pull rows from.</param>
        /// <param name="ct">A token to cancel the enumeration.</param>
        /// <returns>An enumerable that lazily parses and yields one <typeparamref name="T"/> per data row as it is asynchronously enumerated.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public CsvEnumerable<T> ParseAsync(CsvReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new CsvEnumerable<T>(reader, _config, _csvInfo(), ct);
        }
    }

    /// <summary>Creates <see cref="ExcelParser{T}"/> instances.</summary>
    public static class ExcelParser
    {
        /// <summary>
        /// Creates a parser that binds header columns to <typeparamref name="T"/>'s properties decorated with
        /// <c>[ExcelColumn]</c>/<c>[ExcelRequired]</c>/<c>[ExcelConverter]</c>, found by reflection.
        /// </summary>
        /// <remarks>
        /// Not compatible with Native AOT, and trimming can remove the properties it binds to. Use
        /// <see cref="Generated{T}"/> or <see cref="Build{T}"/> instead where that matters. CSV readers get
        /// a separate map with text-based date parsing.
        /// </remarks>
        /// <typeparam name="T">The model type.</typeparam>
        /// <param name="config">Header matching, culture and parse-failure options. Defaults to a new <see cref="ExcelParserConfig"/>.</param>
        /// <returns>The parser.</returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="config"/> has <see cref="ExcelParserConfig.HeaderRow"/> less than 1.</exception>
        [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
        [RequiresDynamicCode("Typed parsing binds property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        public static ExcelParser<T> FromAttributes<T>(ExcelParserConfig? config = null)
            where T : allows ref struct
        {
            return new ExcelParser<T>(ValidateConfig(config), TypeMapper<T>.GetInfo, TypeMapper<T>.GetCsvInfo);
        }

        /// <summary>Creates a parser from the map the <c>[ExcelSerializable]</c> source generator emitted for <typeparamref name="T"/>. Trimming and AOT safe.</summary>
        /// <remarks>
        /// The map is built once and reused for every reader, CSV included. The generator binds
        /// <see cref="DateTime"/>/<see cref="DateOnly"/>/<see cref="TimeOnly"/> properties with the
        /// <c>*Auto</c> readers in <see cref="ExcelCellReaders"/> — serial number first, text as a fallback.
        /// </remarks>
        /// <typeparam name="T">The model type; must implement <see cref="IExcelRowMap{T}"/>.</typeparam>
        /// <param name="config">Header matching, culture and parse-failure options. Defaults to a new <see cref="ExcelParserConfig"/>.</param>
        /// <returns>The parser.</returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="config"/> has <see cref="ExcelParserConfig.HeaderRow"/> less than 1.</exception>
        public static ExcelParser<T> Generated<T>(ExcelParserConfig? config = null)
            where T : IExcelRowMap<T>, allows ref struct
        {
            ExcelParserConfig effective = ValidateConfig(config);
            var builder = new ExcelRowMapBuilder<T>();
            T.ConfigureExcelRowMap(builder);
            return Fixed(effective, builder.Build());
        }

        /// <summary>Creates a parser whose map comes entirely from <paramref name="configure"/>. Trimming and AOT safe.</summary>
        /// <remarks>
        /// For when the mapping is a runtime decision — loaded from config, chosen by a user, or different
        /// per input file. Each call builds its own map, so two parsers configured differently for the same
        /// <typeparamref name="T"/> coexist.
        /// </remarks>
        /// <typeparam name="T">The model type.</typeparam>
        /// <param name="configure">Configures the row map by calling <see cref="ExcelRowMapBuilder{T}.Property{TValue}"/> and its siblings.</param>
        /// <param name="config">Header matching, culture and parse-failure options. Defaults to a new <see cref="ExcelParserConfig"/>.</param>
        /// <returns>The parser.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="configure"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="config"/> has <see cref="ExcelParserConfig.HeaderRow"/> less than 1.</exception>
        public static ExcelParser<T> Build<T>(Action<ExcelRowMapBuilder<T>> configure, ExcelParserConfig? config = null)
            where T : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(configure);
            ExcelParserConfig effective = ValidateConfig(config);
            var builder = new ExcelRowMapBuilder<T>();
            configure(builder);
            return Fixed(effective, builder.Build());
        }

        /// <summary>
        /// Creates a parser whose map merges <paramref name="configure"/>'s bindings with attribute-driven
        /// ones reflected from <typeparamref name="T"/>: a builder binding fully replaces every
        /// attribute-driven property that shares one of its header names, and every other property keeps
        /// its attribute-driven behavior.
        /// </summary>
        /// <remarks>
        /// The match is by header name, not by property identity — the builder only receives a setter
        /// lambda. To override property <c>P</c>, configure the builder with one of the header names
        /// <c>P</c>'s <c>[ExcelColumn]</c> attributes already use. Configuring a <em>different</em> header
        /// name for <c>P</c> overrides nothing: both bindings survive and <c>P</c> is assigned twice per row.
        /// Mark such a property <c>[ExcelIgnore]</c> instead.
        /// </remarks>
        /// <typeparam name="T">The model type.</typeparam>
        /// <param name="configure">Configures the properties that should override their attribute-driven binding.</param>
        /// <param name="config">Header matching, culture and parse-failure options. Defaults to a new <see cref="ExcelParserConfig"/>.</param>
        /// <returns>The parser.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="configure"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="config"/> has <see cref="ExcelParserConfig.HeaderRow"/> less than 1.</exception>
        /// <exception cref="InvalidOperationException"><paramref name="configure"/> calls <see cref="ExcelRowMapBuilder{T}.PropertyAt{TValue}"/>: an index-based map has no header row to match attribute-driven properties against.</exception>
        [RequiresUnreferencedCode("BuildWithAttributeFallback reflects over T's public properties for the attribute-driven fallback, which trimming may remove.")]
        [RequiresDynamicCode("BuildWithAttributeFallback binds attribute-driven property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        public static ExcelParser<T> BuildWithAttributeFallback<T>(Action<ExcelRowMapBuilder<T>> configure, ExcelParserConfig? config = null)
            where T : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(configure);
            ExcelParserConfig effective = ValidateConfig(config);
            var builder = new ExcelRowMapBuilder<T>();
            configure(builder);
            TypeMapInfo<T> fluent = builder.Build(requireFactory: false);
            TypeMapInfo<T> merged = TypeMapInfo<T>.MergeFluentOverAttributes(
                fluent, TypeMapper<T>.GetInfo(), effective.ColumnNameComparer, effective.HeaderNormalization);
            return Fixed(effective, merged);
        }

        private static ExcelParser<T> Fixed<T>(ExcelParserConfig config, TypeMapInfo<T> info)
            where T : allows ref struct
        {
            return new ExcelParser<T>(config, () => info, () => info);
        }

        private static ExcelParserConfig ValidateConfig(ExcelParserConfig? config)
        {
            if (config is not null && config.HeaderRow < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(config), config.HeaderRow, "HeaderRow must be at least 1.");
            }
            return config ?? new ExcelParserConfig();
        }
    }
}
