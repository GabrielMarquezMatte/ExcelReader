namespace ExcelReader.NetFramework.Tests
{
    internal static class Fixtures
    {
        internal static string Path(params string[] parts)
        {
            return System.IO.Path.Combine([AppDomain.CurrentDomain.BaseDirectory, "data", .. parts]);
        }
    }
}
