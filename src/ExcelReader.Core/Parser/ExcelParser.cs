using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser
{
    public sealed class ExcelParser<T> where T : new()
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
        public XlsbExcelEnumerable<T> Parse(XlsbReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new XlsbExcelEnumerable<T>(reader, _config);
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

        public XlsbExcelEnumerable<T> ParseAsync(XlsbReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new XlsbExcelEnumerable<T>(reader, _config, ct);
        }
    }
}
