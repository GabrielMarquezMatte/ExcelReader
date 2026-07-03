namespace ExcelReader.Core.Writer
{
    public sealed record CsvWriterOptions
    {
        public byte Delimiter { get; init; } = (byte)',';
        public byte Quote { get; init; } = (byte)'"';

        public static CsvWriterOptions Default { get; } = new();
    }
}
