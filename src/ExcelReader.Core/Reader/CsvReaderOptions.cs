using System.Text;

namespace ExcelReader.Core.Reader
{
    public sealed record CsvReaderOptions
    {
        public byte Delimiter { get; init; } = (byte)',';
        public byte Quote { get; init; } = (byte)'"';

        // Non-UTF-8 sources (e.g. Windows-1252 exports) are transcoded to UTF-8 at open time so the
        // scanner always works in UTF-8. Null (default) means the source is already UTF-8.
        public Encoding? Encoding { get; init; }
        public bool DetectEncodingFromByteOrderMark { get; init; } = true;

        // Caps a single buffered record/field, mirroring ExcelReaderOptions.MaxCellBytes.
        public int MaxCellBytes { get; init; } = 32 * 1024 * 1024;

        public static CsvReaderOptions Default { get; } = new();
    }
}
