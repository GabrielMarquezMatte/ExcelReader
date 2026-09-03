using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// The reader-dispatch half of a map-driven parser: given a <see cref="TypeMapInfo{T}"/> and an
    /// <see cref="ExcelParserConfig"/>, hands each concrete reader to the enumerable that knows it.
    /// Derived types differ only in where the map comes from — a caller-configured builder
    /// (<see cref="ExcelFluentParser{T}"/>) or <see cref="IExcelRowMap{T}.ConfigureExcelRowMap"/>
    /// (<see cref="ExcelMappedParser{T}"/>).
    /// </summary>
    /// <typeparam name="T">The row model type to bind each row to.</typeparam>
    /// <remarks>
    /// The constructor is <see langword="private protected"/>, so this hierarchy is closed: the two
    /// parsers in this assembly are the only implementations, and the type is public only because a
    /// public sealed class cannot derive from an internal one.
    /// </remarks>
    public abstract class ExcelRowMapParserBase<T>
    {
        private readonly ExcelParserConfig _config;
        private readonly TypeMapInfo<T> _info;

        // Config first so a derived constructor's base(ValidateConfig(config), BuildMap()) rejects a
        // bad HeaderRow before spending anything on the map.
        private protected ExcelRowMapParserBase(ExcelParserConfig config, TypeMapInfo<T> info)
        {
            _config = config;
            _info = info;
        }

        private protected static ExcelParserConfig ValidateConfig(ExcelParserConfig? config)
        {
            if (config is not null && config.HeaderRow < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(config), config.HeaderRow, "HeaderRow must be at least 1.");
            }
            return config ?? new ExcelParserConfig();
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

        /// <summary>Parses the rows of a format-agnostic reader (e.g. one returned by <c>Excel.Open</c>) into <typeparamref name="T"/> instances, lazily as the result is enumerated. Lets callers avoid pattern-matching the concrete reader type; dispatches through the interface enumerator.</summary>
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
        /// <remarks>
        /// Uses a specialized enumerable (dense field binding, single-pass projection) rather than the
        /// generic <see cref="ExcelEnumerable{T}"/>, since CSV rows have no gaps or styles. Prefer this
        /// concrete overload over <see cref="Parse(IExcelRowReader)"/> for CSV — holding the reader as
        /// <see cref="IExcelRowReader"/> instead routes through the generic path.
        /// </remarks>
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
