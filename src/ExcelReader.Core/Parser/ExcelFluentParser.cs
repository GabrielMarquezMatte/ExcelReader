using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Parses rows from an Excel or CSV reader into instances of <typeparamref name="T"/>, using a map
    /// built at runtime by an <see cref="ExcelRowMapBuilder{T}"/> the caller configures — rather than
    /// one derived from <c>[ExcelColumn]</c>/<c>[ExcelRequired]</c> attributes (<see cref="ExcelParser{T}"/>)
    /// or a compile-time map (<see cref="ExcelMappedParser{T}"/>).
    /// </summary>
    /// <typeparam name="T">The row model type to bind each row to.</typeparam>
    /// <remarks>
    /// For when the mapping itself is a runtime decision — loaded from a config file, chosen by a user
    /// in a UI, or different per input file — which no attribute can express. Builds its map once, in
    /// the constructor, from a fresh <see cref="ExcelRowMapBuilder{T}"/> instance; never touches the
    /// static per-type cache <c>TypeMapper{T}</c> uses, so two <see cref="ExcelFluentParser{T}"/>
    /// instances configured differently for the same <typeparamref name="T"/> produce different, correct
    /// results in the same process.
    /// <para>
    /// No <c>[RequiresUnreferencedCode]</c>/<c>[RequiresDynamicCode]</c> here: the constructor's
    /// <c>configure</c> callback is caller-written code wiring hand-picked readers and setters — no
    /// reflection, no <c>MakeGenericMethod</c>/<c>CreateDelegate</c>. This constructor is AOT-clean for
    /// free; <see cref="WithAttributeFallback"/> is the one exception, since it also reflects over
    /// <typeparamref name="T"/> to fall back to attribute-driven properties the builder left unconfigured.
    /// </para>
    /// </remarks>
    public sealed class ExcelFluentParser<T>
    {
        private readonly ExcelParserConfig _config;
        private readonly TypeMapInfo<T> _info;

        /// <summary>Creates a parser whose map comes entirely from <paramref name="configure"/> — no attribute fallback.</summary>
        /// <param name="configure">Configures the row map by calling <see cref="ExcelRowMapBuilder{T}.Property{TValue}"/> and its siblings.</param>
        /// <param name="config">The options controlling header matching, culture, and parse-failure behavior. Defaults to a new <see cref="ExcelParserConfig"/> when <see langword="null"/>.</param>
        /// <exception cref="ArgumentNullException"><paramref name="configure"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="config"/> has <see cref="ExcelParserConfig.HeaderRow"/> less than 1.</exception>
        public ExcelFluentParser(Action<ExcelRowMapBuilder<T>> configure, ExcelParserConfig? config = null)
        {
            ArgumentNullException.ThrowIfNull(configure);
            _config = ValidateConfig(config);
            var builder = new ExcelRowMapBuilder<T>();
            configure(builder);
            _info = builder.Build();
        }

        private ExcelFluentParser(TypeMapInfo<T> info, ExcelParserConfig config)
        {
            _info = info;
            _config = config;
        }

        private static ExcelParserConfig ValidateConfig(ExcelParserConfig? config)
        {
            if (config is not null && config.HeaderRow < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(config), config.HeaderRow, "HeaderRow must be at least 1.");
            }
            return config ?? new ExcelParserConfig();
        }

        /// <summary>
        /// Creates a parser whose map merges <paramref name="configure"/>'s bindings with attribute-driven
        /// ones reflected from <typeparamref name="T"/>: a property the builder configures fully replaces
        /// its attribute (matched by shared header name), and a property the builder never mentions keeps
        /// its attribute-driven behavior. Lets a caller override just the one column that's a runtime
        /// decision without redeclaring the whole model.
        /// </summary>
        /// <param name="configure">Configures the properties that should override their attribute-driven binding.</param>
        /// <param name="config">The options controlling header matching, culture, and parse-failure behavior. Defaults to a new <see cref="ExcelParserConfig"/> when <see langword="null"/>.</param>
        /// <exception cref="ArgumentNullException"><paramref name="configure"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="config"/> has <see cref="ExcelParserConfig.HeaderRow"/> less than 1.</exception>
        [RequiresUnreferencedCode("WithAttributeFallback reflects over T's public properties for the attribute-driven fallback, which trimming may remove.")]
        [RequiresDynamicCode("WithAttributeFallback binds attribute-driven property setters at runtime (MethodInfo.CreateDelegate / MakeGenericMethod).")]
        [SuppressMessage("Design", "CA1000:Do not declare static members on generic types",
            Justification = "Deliberately a second constructor-shaped factory, not a generic utility method: it needs its own [RequiresUnreferencedCode]/[RequiresDynamicCode] pair distinct from the instance constructor's (AOT annotations are static and can't be conditioned on which overload a caller picks), so it can't be a constructor overload.")]
        public static ExcelFluentParser<T> WithAttributeFallback(Action<ExcelRowMapBuilder<T>> configure, ExcelParserConfig? config = null)
        {
            ArgumentNullException.ThrowIfNull(configure);
            ExcelParserConfig effective = ValidateConfig(config);
            var builder = new ExcelRowMapBuilder<T>();
            configure(builder);
            TypeMapInfo<T> fluent = builder.Build();
            TypeMapInfo<T> attributeFallback = TypeMapper<T>.GetInfo();
            TypeMapInfo<T> merged = TypeMapInfo<T>.MergeFluentOverAttributes(fluent, attributeFallback, effective.ColumnNameComparer, effective.HeaderNormalization);
            return new ExcelFluentParser<T>(merged, effective);
        }

        /// <summary>Parses the rows of an XLSX reader into <typeparamref name="T"/> instances, lazily as the result is enumerated.</summary>
        /// <param name="reader">The XLSX reader to pull rows from.</param>
        /// <returns>An enumerable that yields one <typeparamref name="T"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Synchronous entry point; the enumerable also implements IAsyncEnumerable, but ParseAsync is the async counterpart.")]
        public ExcelEnumerable<T> Parse(XlsxReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T>(reader, _config, _info);
        }

        /// <summary>Parses the rows of an XLS reader into <typeparamref name="T"/> instances, lazily as the result is enumerated.</summary>
        /// <param name="reader">The XLS reader to pull rows from.</param>
        /// <returns>An enumerable that yields one <typeparamref name="T"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Synchronous entry point; the enumerable also implements IAsyncEnumerable, but ParseAsync is the async counterpart.")]
        public ExcelEnumerable<T, XlsReader, XlsReader.Enumerator> Parse(XlsReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, XlsReader, XlsReader.Enumerator>(reader, _config, _info);
        }

        /// <summary>Parses the rows of an XLSB reader into <typeparamref name="T"/> instances, lazily as the result is enumerated.</summary>
        /// <param name="reader">The XLSB reader to pull rows from.</param>
        /// <returns>An enumerable that yields one <typeparamref name="T"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Synchronous entry point; the enumerable also implements IAsyncEnumerable, but ParseAsync is the async counterpart.")]
        public ExcelEnumerable<T, XlsbReader, XlsbReader.Enumerator> Parse(XlsbReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, XlsbReader, XlsbReader.Enumerator>(reader, _config, _info);
        }

        /// <summary>Parses the rows of a format-agnostic reader (e.g. one returned by <c>Excel.Open</c>) into <typeparamref name="T"/> instances, lazily as the result is enumerated.</summary>
        /// <param name="reader">The reader to pull rows from.</param>
        /// <returns>An enumerable that yields one <typeparamref name="T"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Synchronous entry point; the enumerable also implements IAsyncEnumerable, but ParseAsync is the async counterpart.")]
        public ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator> Parse(IExcelRowReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator>(reader, _config, _info);
        }

        /// <summary>Parses the rows of a CSV reader into <typeparamref name="T"/> instances, lazily as the result is enumerated.</summary>
        /// <param name="reader">The CSV reader to pull rows from.</param>
        /// <returns>An enumerable that yields one <typeparamref name="T"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Synchronous entry point; the enumerable also implements IAsyncEnumerable, but ParseAsync is the async counterpart.")]
        public CsvEnumerable<T> Parse(CsvReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new CsvEnumerable<T>(reader, _config, _info);
        }

        /// <summary>Parses the rows of an XLSX reader into <typeparamref name="T"/> instances for asynchronous enumeration.</summary>
        /// <param name="reader">The XLSX reader to pull rows from.</param>
        /// <param name="ct">A token to cancel the enumeration.</param>
        /// <returns>An enumerable that lazily parses and yields one <typeparamref name="T"/> per data row as it is asynchronously enumerated.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T> ParseAsync(XlsxReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T>(reader, _config, _info, ct);
        }

        /// <summary>Parses the rows of an XLS reader into <typeparamref name="T"/> instances for asynchronous enumeration.</summary>
        /// <param name="reader">The XLS reader to pull rows from.</param>
        /// <param name="ct">A token to cancel the enumeration.</param>
        /// <returns>An enumerable that lazily parses and yields one <typeparamref name="T"/> per data row as it is asynchronously enumerated.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T, XlsReader, XlsReader.Enumerator> ParseAsync(XlsReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, XlsReader, XlsReader.Enumerator>(reader, _config, _info, ct);
        }

        /// <summary>Parses the rows of an XLSB reader into <typeparamref name="T"/> instances for asynchronous enumeration.</summary>
        /// <param name="reader">The XLSB reader to pull rows from.</param>
        /// <param name="ct">A token to cancel the enumeration.</param>
        /// <returns>An enumerable that lazily parses and yields one <typeparamref name="T"/> per data row as it is asynchronously enumerated.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T, XlsbReader, XlsbReader.Enumerator> ParseAsync(XlsbReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, XlsbReader, XlsbReader.Enumerator>(reader, _config, _info, ct);
        }

        /// <summary>Parses the rows of a format-agnostic reader (e.g. one returned by <c>Excel.Open</c>) into <typeparamref name="T"/> instances for asynchronous enumeration.</summary>
        /// <param name="reader">The reader to pull rows from.</param>
        /// <param name="ct">A token to cancel the enumeration.</param>
        /// <returns>An enumerable that lazily parses and yields one <typeparamref name="T"/> per data row as it is asynchronously enumerated.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator> ParseAsync(IExcelRowReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator>(reader, _config, _info, ct);
        }

        /// <summary>Parses the rows of a CSV reader into <typeparamref name="T"/> instances for asynchronous enumeration.</summary>
        /// <param name="reader">The CSV reader to pull rows from.</param>
        /// <param name="ct">A token to cancel the enumeration.</param>
        /// <returns>An enumerable that lazily parses and yields one <typeparamref name="T"/> per data row as it is asynchronously enumerated.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public CsvEnumerable<T> ParseAsync(CsvReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new CsvEnumerable<T>(reader, _config, _info, ct);
        }
    }
}
