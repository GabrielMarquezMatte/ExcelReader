namespace ExcelReader.Core.Reader
{
    public static class Excel
    {
        public static XlsxReader FromFile(string path)
        {
            return new XlsxReader(File.OpenRead(path), leaveOpen: false);
        }

        public static XlsxReader From(Stream stream, bool leaveOpen = true)
        {
            return new XlsxReader(stream, leaveOpen);
        }
    }
}
