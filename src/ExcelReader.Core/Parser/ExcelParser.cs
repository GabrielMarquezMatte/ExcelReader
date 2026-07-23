using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Parses rows from an Excel or CSV reader into instances of <typeparamref name="T"/> by matching
    /// header columns to properties decorated with <c>[ExcelColumn]</c>/<c>[ExcelRequired]</c>/<c>[ExcelConverter]</c>.
    /// </summary>
    /// <typeparam name="T">The model type to bind each row to.</typeparam>
    /// <remarks>
    /// Reflects over <typeparamref name="T"/>'s properties (<c>GetProperties</c>) and compiles
    /// per-property setters via <c>Expression.Compile</c> plus <c>MakeGenericMethod</c>, so it needs
    /// runtime code generation and keeps <typeparamref name="T"/>'s members — not compatible with
    /// Native AOT, and trimming can remove the properties it binds to. The raw <c>Excel.From*</c>
    /// readers use no reflection and stay AOT/trim-safe; only this typed layer does not.
    /// <para>
    /// Column binding runs through <c>ref TModel</c> end to end (<c>ColumnParser&lt;T&gt;</c>,
    /// <c>RefAction&lt;T,TProperty&gt;</c> — see <c>Internal/Delegates.cs</c>), and <c>Row</c>/<c>RowCell</c>
    /// are ref structs, so a <c>struct T</c> consumed via a direct <c>foreach</c> (not LINQ over
    /// <c>IEnumerable&lt;object&gt;</c> or anything else that boxes) skips the per-row model allocation a
    /// class <typeparamref name="T"/> requires — measured -59% (3.88 MB -&gt; 1.59 MB / 50k rows) on a
    /// 4-column benchmark record. The rest of that allocation is <typeparamref name="T"/>'s own
    /// reference-typed fields (e.g. a string column decodes to a fresh managed string per row regardless
    /// of <typeparamref name="T"/>'s kind) — struct <typeparamref name="T"/> doesn't remove that, only
    /// the container. See <c>ParseBenchmark.ExcelParserStructSync/RecordStruct</c> in
    /// <c>tests/ExcelReader.Benchmarks</c>.
    /// </para>
    /// </remarks>
    [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
    [RequiresDynamicCode("Typed parsing compiles property setters at runtime (Expression.Compile / MakeGenericMethod).")]
    public sealed class ExcelParser<T>
    {
        private readonly ExcelParserConfig _config;

        /// <summary>Creates a parser configured with the given options, or with defaults if none are supplied.</summary>
        /// <param name="config">The options controlling header matching, culture, and parse-failure behavior. Defaults to a new <see cref="ExcelParserConfig"/> when <see langword="null"/>.</param>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="config"/> has <see cref="ExcelParserConfig.HeaderRow"/> less than 1.</exception>
        public ExcelParser(ExcelParserConfig? config = null)
        {
            if (config is not null && config.HeaderRow < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(config), config.HeaderRow, "HeaderRow must be at least 1.");
            }
            _config = config ?? new ExcelParserConfig();
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
            return new ExcelEnumerable<T>(reader, _config);
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
            return new ExcelEnumerable<T, XlsReader, XlsReader.Enumerator>(reader, _config);
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
            return new ExcelEnumerable<T, XlsbReader, XlsbReader.Enumerator>(reader, _config);
        }

        /// <summary>Parses the rows of a format-agnostic reader (e.g. one returned by <c>Excel.Open</c>) into <typeparamref name="T"/> instances, lazily as the result is enumerated. Lets callers avoid pattern-matching the concrete reader type; dispatches through the interface enumerator.</summary>
        /// <param name="reader">The reader to pull rows from.</param>
        /// <returns>An enumerable that yields one <typeparamref name="T"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Synchronous entry point; the enumerable also implements IAsyncEnumerable, but ParseAsync is the async counterpart.")]
        public ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator> Parse(IExcelRowReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator>(reader, _config);
        }

        /// <summary>Parses the rows of a CSV reader into <typeparamref name="T"/> instances, lazily as the result is enumerated.</summary>
        /// <param name="reader">The CSV reader to pull rows from.</param>
        /// <returns>An enumerable that yields one <typeparamref name="T"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        /// <remarks>
        /// Uses a specialized enumerable (dense field binding, single-pass projection, native text date
        /// parsing) rather than the generic <see cref="ExcelEnumerable{T}"/>, since CSV rows have no
        /// gaps, styles, or serial dates. Prefer this concrete overload over <see cref="Parse(IExcelRowReader)"/>
        /// for CSV — holding the reader as <see cref="IExcelRowReader"/> instead routes through the
        /// generic path (serial-date semantics).
        /// </remarks>
        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Synchronous entry point; the enumerable also implements IAsyncEnumerable, but ParseAsync is the async counterpart.")]
        public CsvEnumerable<T> Parse(CsvReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new CsvEnumerable<T>(reader, _config);
        }

        /// <summary>Parses the rows of an XLSX reader into <typeparamref name="T"/> instances for asynchronous enumeration.</summary>
        /// <param name="reader">The XLSX reader to pull rows from.</param>
        /// <param name="ct">A token to cancel the enumeration.</param>
        /// <returns>An enumerable that lazily parses and yields one <typeparamref name="T"/> per data row as it is asynchronously enumerated.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T> ParseAsync(XlsxReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T>(reader, _config, ct);
        }

        /// <summary>Parses the rows of an XLS reader into <typeparamref name="T"/> instances for asynchronous enumeration.</summary>
        /// <param name="reader">The XLS reader to pull rows from.</param>
        /// <param name="ct">A token to cancel the enumeration.</param>
        /// <returns>An enumerable that lazily parses and yields one <typeparamref name="T"/> per data row as it is asynchronously enumerated.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T, XlsReader, XlsReader.Enumerator> ParseAsync(XlsReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, XlsReader, XlsReader.Enumerator>(reader, _config, ct);
        }

        /// <summary>Parses the rows of an XLSB reader into <typeparamref name="T"/> instances for asynchronous enumeration.</summary>
        /// <param name="reader">The XLSB reader to pull rows from.</param>
        /// <param name="ct">A token to cancel the enumeration.</param>
        /// <returns>An enumerable that lazily parses and yields one <typeparamref name="T"/> per data row as it is asynchronously enumerated.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T, XlsbReader, XlsbReader.Enumerator> ParseAsync(XlsbReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, XlsbReader, XlsbReader.Enumerator>(reader, _config, ct);
        }

        /// <summary>Parses the rows of a format-agnostic reader (e.g. one returned by <c>Excel.Open</c>) into <typeparamref name="T"/> instances for asynchronous enumeration.</summary>
        /// <param name="reader">The reader to pull rows from.</param>
        /// <param name="ct">A token to cancel the enumeration.</param>
        /// <returns>An enumerable that lazily parses and yields one <typeparamref name="T"/> per data row as it is asynchronously enumerated.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator> ParseAsync(IExcelRowReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator>(reader, _config, ct);
        }

        /// <summary>Parses the rows of a CSV reader into <typeparamref name="T"/> instances for asynchronous enumeration.</summary>
        /// <param name="reader">The CSV reader to pull rows from.</param>
        /// <param name="ct">A token to cancel the enumeration.</param>
        /// <returns>An enumerable that lazily parses and yields one <typeparamref name="T"/> per data row as it is asynchronously enumerated.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        public CsvEnumerable<T> ParseAsync(CsvReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new CsvEnumerable<T>(reader, _config, ct);
        }
    }
}
