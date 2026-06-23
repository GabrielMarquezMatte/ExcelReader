namespace ExcelReader.Core.Parser
{
    public sealed class ExcelParserConfig
    {
        public int HeaderRow { get; init; } = 1;
        public StringComparer ColumnNameComparer { get; init; } = StringComparer.OrdinalIgnoreCase;
    }
}
