using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Parses rows from an Excel or CSV reader into instances of <typeparamref name="T"/>, exactly like
    /// <see cref="ExcelParser{T}"/>, but from a map <see cref="IExcelRowMap{T}.ConfigureExcelRowMap"/>
    /// builds (source-generated or hand-written) instead of one built by reflecting over <typeparamref name="T"/>.
    /// </summary>
    /// <typeparam name="T">
    /// The model type to bind each row to; must implement <see cref="IExcelRowMap{T}"/>. Not implementing
    /// it is a compile error here, rather than a silent runtime fallback to reflection.
    /// </typeparam>
    /// <remarks>
    /// No <c>[RequiresUnreferencedCode]</c>/<c>[RequiresDynamicCode]</c>: the <c>where T : IExcelRowMap&lt;T&gt;</c>
    /// constraint guarantees <see cref="TypeMapInfo{T}"/> comes from <typeparamref name="T"/>'s own
    /// <see cref="IExcelRowMap{T}.ConfigureExcelRowMap"/>, which reaches neither <c>GetProperties</c> nor
    /// <c>MakeGenericMethod</c>/<c>CreateDelegate</c> — so nothing here needs those annotations, and
    /// nothing here can regress into needing them.
    /// <para>
    /// Unlike <see cref="ExcelParser{T}"/>, which builds a separate CSV-specific map with text-based date
    /// parsing (<c>TypeMapper&lt;T&gt;.GetCsvInfo()</c>), this type builds <typeparamref name="T"/>'s map
    /// exactly once and reuses it for every reader, including <see cref="Parse(CsvReader)"/>. A
    /// model with a date property meant to be read from CSV should bind it with
    /// <see cref="ExcelCellReaders.DateTimeText"/>/<see cref="ExcelCellReaders.DateOnlyText"/>/<see cref="ExcelCellReaders.TimeOnlyText"/>
    /// rather than the serial-number readers.
    /// </para>
    /// </remarks>
    public sealed class ExcelMappedParser<T> where T : IExcelRowMap<T>
    {
        private readonly ExcelParserConfig _config;
        private readonly TypeMapInfo<T> _info;

        /// <summary>Creates a parser configured with the given options, or with defaults if none are supplied.</summary>
        /// <param name="config">The options controlling header matching, culture, and parse-failure behavior. Defaults to a new <see cref="ExcelParserConfig"/> when <see langword="null"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="config"/> has <see cref="ExcelParserConfig.HeaderRow"/> less than 1.</exception>
        public ExcelMappedParser(ExcelParserConfig? config = null)
        {
            if (config is not null && config.HeaderRow < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(config), config.HeaderRow, "HeaderRow must be at least 1.");
            }
            _config = config ?? new ExcelParserConfig();
            var builder = new ExcelRowMapBuilder<T>();
            T.ConfigureExcelRowMap(builder);
            _info = builder.Build();
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
