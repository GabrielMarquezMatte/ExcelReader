namespace ExcelReader.NetFramework.Tests
{
    internal static class Fixtures
    {
        internal static string Path(params string[] parts)
        {
            string root = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data");
            return System.IO.Path.Combine(root, System.IO.Path.Combine(parts));
        }
    }
}
