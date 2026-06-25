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

        public ExcelEnumerable<T> Parse(XlsxReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelEnumerable<T>(reader, _config);
        }

        public XlsExcelEnumerable<T> Parse(XlsReader reader)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new XlsExcelEnumerable<T>(reader, _config);
        }

        public ExcelAsyncEnumerable<T> ParseAsync(XlsxReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new ExcelAsyncEnumerable<T>(reader, _config, ct);
        }

        public XlsExcelAsyncEnumerable<T> ParseAsync(XlsReader reader, CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(reader);
            return new XlsExcelAsyncEnumerable<T>(reader, _config, ct);
        }
    }
}
