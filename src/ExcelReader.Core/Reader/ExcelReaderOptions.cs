namespace ExcelReader.Core.Reader
{
    // Options are an explicit extension point — empty for now (no settings to configure yet).
    public readonly record struct ExcelReaderOptions
    {
        public XlsxReader FromFile(string path)
        {
            return new(File.OpenRead(path), leaveOpen: false);
        }

        public XlsxReader From(Stream stream, bool leaveOpen = true)
        {
            return new(stream, leaveOpen);
        }
    }
}
