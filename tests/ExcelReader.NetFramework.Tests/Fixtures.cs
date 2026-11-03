namespace ExcelReader.NetFramework.Tests
{
    internal static class Fixtures
    {
        internal static string Path(params string[] parts)
            => System.IO.Path.Combine(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data"), System.IO.Path.Combine(parts));
    }
}
