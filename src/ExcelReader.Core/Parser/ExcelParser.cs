using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser
{
    // Typed parsing reflects over T's properties (GetProperties) and compiles per-property setters via
    // Expression.Compile plus MakeGenericMethod, so it needs runtime code generation and keeps T's
    // members. Not compatible with Native AOT, and trimming can remove the properties it binds to. The
    // raw Excel.From* readers use no reflection and stay AOT/trim-safe; only this typed layer does not.
    [RequiresUnreferencedCode("Typed parsing reflects over T's public properties, which trimming may remove.")]
    [RequiresDynamicCode("Typed parsing compiles property setters at runtime (Expression.Compile / MakeGenericMethod).")]
    public sealed class ExcelParser<T>
    {
        private readonly ExcelParserConfig _config;

        public ExcelParser(ExcelParserConfig? config = null)
        {
            if (config is not null && config.HeaderRow < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(config), config.HeaderRow, "HeaderRow must be at least 1.");
            }
            _config = config ?? new ExcelParserConfig();
        }

        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Synchronous entry point; the enumerable also implements IAsyncEnumerable, but ParseAsync is the async counterpart.")]
        public ExcelEnumerable<T> Parse(XlsxReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T>(reader, _config);
        }

        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Synchronous entry point; the enumerable also implements IAsyncEnumerable, but ParseAsync is the async counterpart.")]
        public XlsExcelEnumerable<T> Parse(XlsReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new XlsExcelEnumerable<T>(reader, _config);
        }

        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Synchronous entry point; the enumerable also implements IAsyncEnumerable, but ParseAsync is the async counterpart.")]
        public ExcelEnumerable<T, XlsbReader, XlsbReader.Enumerator> Parse(XlsbReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, XlsbReader, XlsbReader.Enumerator>(reader, _config);
        }

        // Format-agnostic entry point for the reader returned by Excel.Open, so callers need not
        // pattern-match the concrete reader type. Dispatches through the interface enumerator.
        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Synchronous entry point; the enumerable also implements IAsyncEnumerable, but ParseAsync is the async counterpart.")]
        public ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator> Parse(IExcelRowReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator>(reader, _config);
        }

        // CSV uses a specialized enumerable (dense field binding, single-pass projection, native text
        // date parsing) rather than the generic ExcelEnumerable, since CSV rows have no gaps, styles,
        // or serial dates. Holding the reader as IExcelRowReader instead routes through the generic
        // path (serial-date semantics), so prefer this concrete overload for CSV.
        [SuppressMessage("Usage", "VSTHRD200:Use \"Async\" suffix for async methods",
            Justification = "Synchronous entry point; the enumerable also implements IAsyncEnumerable, but ParseAsync is the async counterpart.")]
        public CsvEnumerable<T> Parse(CsvReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new CsvEnumerable<T>(reader, _config);
        }

        public ExcelEnumerable<T> ParseAsync(XlsxReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T>(reader, _config, ct);
        }

        public XlsExcelEnumerable<T> ParseAsync(XlsReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new XlsExcelEnumerable<T>(reader, _config, ct);
        }

        public ExcelEnumerable<T, XlsbReader, XlsbReader.Enumerator> ParseAsync(XlsbReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, XlsbReader, XlsbReader.Enumerator>(reader, _config, ct);
        }

        public ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator> ParseAsync(IExcelRowReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T, IExcelRowReader, IExcelRowEnumerator>(reader, _config, ct);
        }

        public CsvEnumerable<T> ParseAsync(CsvReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new CsvEnumerable<T>(reader, _config, ct);
        }
    }
}
