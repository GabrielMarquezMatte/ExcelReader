using System.Reflection;
using System.Runtime.ExceptionServices;

namespace ExcelReader.Core.Parser.Internal
{
    internal static class TypeMapper<T> where T : new()
    {
        private static readonly Lazy<TypeMapInfo<T>> _info =
            new(BuildSafe, LazyThreadSafetyMode.ExecutionAndPublication);

        internal static TypeMapInfo<T> GetInfo()
        {
            return _info.Value;
        }

        private static TypeMapInfo<T> BuildSafe()
        {
            try
            {
                return Build();
            }
            catch (Exception ex)
            {
                // Capture here so the exception is re-thrown with original stack trace on every call.
                ExceptionDispatchInfo.Capture(ex).Throw();
                throw; // unreachable; satisfies compiler
            }
        }

        private static TypeMapInfo<T> Build()
        {
            PropertyInfo[] properties = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance);

            var names = new List<string>(properties.Length);
            var parsers = new List<ColumnParser<T>>(properties.Length);

            foreach (PropertyInfo prop in properties)
            {
                if (!prop.CanWrite || prop.GetSetMethod() is null)
                {
                    continue;
                }
                ColumnParser<T>? parser = ColumnParserFactory.Build<T>(prop);
                if (parser is null)
                {
                    continue;
                }
                ExcelColumnAttribute? attr = prop.GetCustomAttribute<ExcelColumnAttribute>();
                string name = attr is not null ? attr.Name : prop.Name;
                names.Add(name);
                parsers.Add(parser);
            }

            return new TypeMapInfo<T>([.. names], [.. parsers]);
        }
    }

    internal readonly struct TypeMapInfo<T> where T : new()
    {
        internal readonly string[] Names;
        internal readonly ColumnParser<T>[] Parsers;

        internal TypeMapInfo(string[] names, ColumnParser<T>[] parsers)
        {
            Names = names;
            Parsers = parsers;
        }

        // Linear scan — property counts are small; avoids allocating a dictionary per config.
        internal int FindIndex(string headerName, StringComparer comparer)
        {
            for (int i = 0; i < Names.Length; i++)
            {
                if (comparer.Equals(Names[i], headerName))
                {
                    return i;
                }
            }
            return -1;
        }
    }
}
