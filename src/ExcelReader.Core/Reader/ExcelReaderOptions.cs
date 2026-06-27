namespace ExcelReader.Core.Reader
{
    public sealed record ExcelReaderOptions
    {
        public long MaxTotalDecompressedBytes { get; init; } = 512L * 1024 * 1024;
        public int MaxCellBytes { get; init; } = 32 * 1024 * 1024;
        public long MaxSharedStringBytes { get; init; } = 128L * 1024 * 1024;

        public static ExcelReaderOptions Default { get; } = new();
    }
}
