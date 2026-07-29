#if NET9_0_OR_GREATER
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Parses rows into instances of a <c>ref struct</c>-eligible model type by matching header columns to
    /// properties decorated with <c>[ExcelColumn]</c>/<c>[ExcelRequired]</c>/<c>[ExcelConverter]</c>, the same
    /// way <see cref="ExcelParser{T}"/> does for ordinary types.
    /// </summary>
    /// <remarks>
    /// Covers the one shape <see cref="ExcelParser{T}"/> cannot target — a ref struct can't be a type
    /// argument to <c>IEnumerable&lt;T&gt;</c>/<c>IAsyncEnumerable&lt;T&gt;</c> before this TFM. Uses the
    /// same <c>TypeMapper&lt;T&gt;</c>/<c>ColumnParserFactory</c> machinery as <see cref="ExcelParser{T}"/>,
    /// widened to <c>allows ref struct</c>. Not AOT/trim-safe for the same reason
    /// <see cref="ExcelParser{T}"/> isn't (reflection + <c>Expression.Compile</c> at type-map-build
    /// time, cached per closed <c>TModel</c> thereafter).
    /// <para>
    /// Sync-only, permanently: <c>IAsyncEnumerable&lt;TModel&gt;</c> cannot have a ref struct element type
    /// (CS9267), so there is no async counterpart here, ever — not merely unimplemented. Callers who need
    /// async should keep <see cref="ExcelParser{T}"/> with a plain struct/class model.
    /// </para>
    /// <para>
    /// A <c>ReadOnlySpan&lt;byte&gt;</c> property binds directly to <c>Cell.Value</c> — zero-copy, aliasing
    /// the reader's row/shared-string buffer, valid only until the next <c>MoveNext()</c> (copy it out,
    /// e.g. via <c>Encoding.UTF8.GetString(span)</c>, to retain it past the loop body). A <c>string</c>
    /// property still works but allocates per row (same as <see cref="ExcelParser{T}"/>). Any other
    /// unsupported property type is silently left unbound (matching <see cref="ExcelParser{T}"/>'s
    /// existing behavior) unless marked <c>[ExcelRequired]</c>, which throws at type-map-build time
    /// instead. A model using only <c>ReadOnlySpan&lt;byte&gt;</c>/numeric/bool/date columns is fully
    /// zero-alloc end to end — not just the container. A plain <c>struct</c> model with
    /// <see cref="ExcelParser{T}"/> already avoids allocating the container itself; this closes the
    /// remaining gap by letting a text column bind to a span instead of allocating a <c>string</c> too.
    /// </para>
    /// </remarks>
    public static class RefParser
    {
        /// <summary>Parses the rows of an XLSX reader into <typeparamref name="TModel"/> instances, lazily as the result is enumerated.</summary>
        /// <param name="reader">The XLSX reader to pull rows from.</param>
        /// <param name="config">The options controlling header matching, culture, and parse-failure behavior. Defaults to a new <see cref="ExcelParserConfig"/> when <see langword="null"/>.</param>
        /// <returns>An enumerable that yields one <typeparamref name="TModel"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        [RequiresUnreferencedCode("ParseNamed reflects over TModel's public properties, which trimming may remove.")]
        [RequiresDynamicCode("ParseNamed compiles property setters at runtime (Expression.Compile / MakeGenericMethod).")]
        public static NamedRefRowEnumerable<TModel, XlsxReader, XlsxReader.Enumerator> ParseNamed<TModel>(
            XlsxReader reader, ExcelParserConfig? config = null)
            where TModel : allows ref struct
        {
            return Create<TModel, XlsxReader, XlsxReader.Enumerator>(reader, TypeMapper<TModel>.GetInfo(), config);
        }

        /// <summary>Parses the rows of an XLSB reader into <typeparamref name="TModel"/> instances, lazily as the result is enumerated.</summary>
        /// <param name="reader">The XLSB reader to pull rows from.</param>
        /// <param name="config">The options controlling header matching, culture, and parse-failure behavior. Defaults to a new <see cref="ExcelParserConfig"/> when <see langword="null"/>.</param>
        /// <returns>An enumerable that yields one <typeparamref name="TModel"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        [RequiresUnreferencedCode("ParseNamed reflects over TModel's public properties, which trimming may remove.")]
        [RequiresDynamicCode("ParseNamed compiles property setters at runtime (Expression.Compile / MakeGenericMethod).")]
        public static NamedRefRowEnumerable<TModel, XlsbReader, XlsbReader.Enumerator> ParseNamed<TModel>(
            XlsbReader reader, ExcelParserConfig? config = null)
            where TModel : allows ref struct
        {
            return Create<TModel, XlsbReader, XlsbReader.Enumerator>(reader, TypeMapper<TModel>.GetInfo(), config);
        }

        /// <summary>Parses the rows of an XLS reader into <typeparamref name="TModel"/> instances, lazily as the result is enumerated.</summary>
        /// <param name="reader">The XLS reader to pull rows from.</param>
        /// <param name="config">The options controlling header matching, culture, and parse-failure behavior. Defaults to a new <see cref="ExcelParserConfig"/> when <see langword="null"/>.</param>
        /// <returns>An enumerable that yields one <typeparamref name="TModel"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        [RequiresUnreferencedCode("ParseNamed reflects over TModel's public properties, which trimming may remove.")]
        [RequiresDynamicCode("ParseNamed compiles property setters at runtime (Expression.Compile / MakeGenericMethod).")]
        public static NamedRefRowEnumerable<TModel, XlsReader, XlsReader.Enumerator> ParseNamed<TModel>(
            XlsReader reader, ExcelParserConfig? config = null)
            where TModel : allows ref struct
        {
            return Create<TModel, XlsReader, XlsReader.Enumerator>(reader, TypeMapper<TModel>.GetInfo(), config);
        }

        /// <summary>Parses the rows of a CSV reader into <typeparamref name="TModel"/> instances, lazily as the result is enumerated.</summary>
        /// <param name="reader">The CSV reader to pull rows from.</param>
        /// <param name="config">The options controlling header matching, culture, and parse-failure behavior. Defaults to a new <see cref="ExcelParserConfig"/> when <see langword="null"/>.</param>
        /// <returns>An enumerable that yields one <typeparamref name="TModel"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        /// <remarks>
        /// Uses <c>TypeMapper&lt;TModel&gt;.GetCsvInfo()</c> — the only difference from the other overloads
        /// is that <c>DateTime</c>/<c>DateOnly</c>/<c>TimeOnly</c> columns parse the cell text rather than
        /// an Excel serial number (see <c>TypeMapper.GetCsvInfo</c>).
        /// </remarks>
        [RequiresUnreferencedCode("ParseNamed reflects over TModel's public properties, which trimming may remove.")]
        [RequiresDynamicCode("ParseNamed compiles property setters at runtime (Expression.Compile / MakeGenericMethod).")]
        public static NamedRefRowEnumerable<TModel, CsvReader, CsvReader.Enumerator> ParseNamed<TModel>(
            CsvReader reader, ExcelParserConfig? config = null)
            where TModel : allows ref struct
        {
            return Create<TModel, CsvReader, CsvReader.Enumerator>(reader, TypeMapper<TModel>.GetCsvInfo(), config);
        }

        /// <summary>Parses the rows of a format-agnostic reader (e.g. one returned by <c>Excel.Open</c>) into <typeparamref name="TModel"/> instances, lazily as the result is enumerated. Lets callers avoid pattern-matching the concrete reader type; dispatches through the interface enumerator.</summary>
        /// <param name="reader">The reader to pull rows from.</param>
        /// <param name="config">The options controlling header matching, culture, and parse-failure behavior. Defaults to a new <see cref="ExcelParserConfig"/> when <see langword="null"/>.</param>
        /// <returns>An enumerable that yields one <typeparamref name="TModel"/> per data row.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="reader"/> is <see langword="null"/>.</exception>
        [RequiresUnreferencedCode("ParseNamed reflects over TModel's public properties, which trimming may remove.")]
        [RequiresDynamicCode("ParseNamed compiles property setters at runtime (Expression.Compile / MakeGenericMethod).")]
        public static NamedRefRowEnumerable<TModel, IExcelRowReader, IExcelRowEnumerator> ParseNamed<TModel>(
            IExcelRowReader reader, ExcelParserConfig? config = null)
            where TModel : allows ref struct
        {
            return Create<TModel, IExcelRowReader, IExcelRowEnumerator>(reader, TypeMapper<TModel>.GetInfo(), config);
        }

        private static NamedRefRowEnumerable<TModel, TReader, TEnumerator> Create<TModel, TReader, TEnumerator>(
            TReader reader, TypeMapInfo<TModel> typeInfo, ExcelParserConfig? config)
            where TModel : allows ref struct
            where TReader : class, IExcelRowReader<TEnumerator>
            where TEnumerator : class, IExcelRowEnumerator
        {
            ArgumentNullException.ThrowIfNull(reader);
            config ??= new ExcelParserConfig();
            return new NamedRefRowEnumerable<TModel, TReader, TEnumerator>(
                reader, typeInfo, config.ColumnNameComparer, config.HeaderNormalization, config.HeaderRow, config.Culture, config.ThrowOnParseFailure);
        }
    }
}
#endif
