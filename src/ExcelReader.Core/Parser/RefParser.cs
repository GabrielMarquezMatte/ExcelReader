#if NET9_0_OR_GREATER
using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser
{
    // Reflection-based, attribute-driven typed parsing for models that may be a `ref struct` — the
    // one shape ExcelParser<T> cannot target (a ref struct can't be a type argument to
    // IEnumerable<T>/IAsyncEnumerable<T> before this TFM). Properties are matched to header columns by
    // name via [ExcelColumn]/[ExcelRequired]/[ExcelConverter], exactly like ExcelParser<T> — same
    // TypeMapper<T>/ColumnParserFactory machinery underneath, widened to `allows ref struct` for this
    // TFM. Not AOT/trim-safe, for the same reason ExcelParser<T> isn't (reflection + Expression.Compile
    // at type-map-build time, cached per closed TModel thereafter).
    //
    // Sync-only, permanently: IAsyncEnumerable<TModel> cannot have a ref struct element type (CS9267),
    // so there is no ParseAsync here, ever — not merely unimplemented. Callers who need async keep
    // ExcelParser<T> with a plain struct/class model.
    //
    // Text columns: a `ReadOnlySpan<byte>` property binds directly to Cell.Value — zero-copy, aliases
    // the reader's row/shared-string buffer, valid only until the next MoveNext() (copy it out, e.g.
    // via Encoding.UTF8.GetString(span), to retain it past the loop body). A `string` property still
    // works too but allocates per row (same as ExcelParser<T>). Any other unsupported property type is
    // silently left unbound (matching ExcelParser<T>'s existing behavior) unless marked [ExcelRequired],
    // which throws at type-map-build time instead. A model using only ReadOnlySpan<byte>/numeric/bool/
    // date columns is fully zero-alloc end to end — not just the container (see docs/performance-plan.md
    // P2, which measured the container-only win for a plain struct; this closes the remaining gap for
    // genuine ref structs too).
    public static class RefParser
    {
        [RequiresUnreferencedCode("ParseNamed reflects over TModel's public properties, which trimming may remove.")]
        [RequiresDynamicCode("ParseNamed compiles property setters at runtime (Expression.Compile / MakeGenericMethod).")]
        public static NamedRefRowEnumerable<TModel, XlsxReader, XlsxReader.Enumerator> ParseNamed<TModel>(
            XlsxReader reader, ExcelParserConfig? config = null)
            where TModel : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(reader);
            config ??= new ExcelParserConfig();
            return new NamedRefRowEnumerable<TModel, XlsxReader, XlsxReader.Enumerator>(
                reader, TypeMapper<TModel>.GetInfo(), config.ColumnNameComparer, config.HeaderNormalization, config.HeaderRow, config.Culture, config.ThrowOnParseFailure);
        }

        [RequiresUnreferencedCode("ParseNamed reflects over TModel's public properties, which trimming may remove.")]
        [RequiresDynamicCode("ParseNamed compiles property setters at runtime (Expression.Compile / MakeGenericMethod).")]
        public static NamedRefRowEnumerable<TModel, XlsbReader, XlsbReader.Enumerator> ParseNamed<TModel>(
            XlsbReader reader, ExcelParserConfig? config = null)
            where TModel : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(reader);
            config ??= new ExcelParserConfig();
            return new NamedRefRowEnumerable<TModel, XlsbReader, XlsbReader.Enumerator>(
                reader, TypeMapper<TModel>.GetInfo(), config.ColumnNameComparer, config.HeaderNormalization, config.HeaderRow, config.Culture, config.ThrowOnParseFailure);
        }

        [RequiresUnreferencedCode("ParseNamed reflects over TModel's public properties, which trimming may remove.")]
        [RequiresDynamicCode("ParseNamed compiles property setters at runtime (Expression.Compile / MakeGenericMethod).")]
        public static NamedRefRowEnumerable<TModel, XlsReader, XlsReader.Enumerator> ParseNamed<TModel>(
            XlsReader reader, ExcelParserConfig? config = null)
            where TModel : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(reader);
            config ??= new ExcelParserConfig();
            return new NamedRefRowEnumerable<TModel, XlsReader, XlsReader.Enumerator>(
                reader, TypeMapper<TModel>.GetInfo(), config.ColumnNameComparer, config.HeaderNormalization, config.HeaderRow, config.Culture, config.ThrowOnParseFailure);
        }

        // CSV uses TypeMapper<TModel>.GetCsvInfo() — the only difference is DateTime/DateOnly/TimeOnly
        // columns parse the cell text rather than an Excel serial number (see TypeMapper.GetCsvInfo).
        [RequiresUnreferencedCode("ParseNamed reflects over TModel's public properties, which trimming may remove.")]
        [RequiresDynamicCode("ParseNamed compiles property setters at runtime (Expression.Compile / MakeGenericMethod).")]
        public static NamedRefRowEnumerable<TModel, CsvReader, CsvReader.Enumerator> ParseNamed<TModel>(
            CsvReader reader, ExcelParserConfig? config = null)
            where TModel : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(reader);
            config ??= new ExcelParserConfig();
            return new NamedRefRowEnumerable<TModel, CsvReader, CsvReader.Enumerator>(
                reader, TypeMapper<TModel>.GetCsvInfo(), config.ColumnNameComparer, config.HeaderNormalization, config.HeaderRow, config.Culture, config.ThrowOnParseFailure);
        }

        // Format-agnostic entry point for the reader returned by Excel.Open, so callers need not
        // pattern-match the concrete reader type. Dispatches through the interface enumerator.
        [RequiresUnreferencedCode("ParseNamed reflects over TModel's public properties, which trimming may remove.")]
        [RequiresDynamicCode("ParseNamed compiles property setters at runtime (Expression.Compile / MakeGenericMethod).")]
        public static NamedRefRowEnumerable<TModel, IExcelRowReader, IExcelRowEnumerator> ParseNamed<TModel>(
            IExcelRowReader reader, ExcelParserConfig? config = null)
            where TModel : allows ref struct
        {
            ArgumentNullException.ThrowIfNull(reader);
            config ??= new ExcelParserConfig();
            return new NamedRefRowEnumerable<TModel, IExcelRowReader, IExcelRowEnumerator>(
                reader, TypeMapper<TModel>.GetInfo(), config.ColumnNameComparer, config.HeaderNormalization, config.HeaderRow, config.Culture, config.ThrowOnParseFailure);
        }
    }
}
#endif
