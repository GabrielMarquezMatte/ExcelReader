namespace ExcelReader.Core.Writer.Internal
{
    internal sealed class TableRegistry
    {
        private readonly HashSet<string> _names = new(StringComparer.OrdinalIgnoreCase);
        private int _nextId = 1;

        internal void RequireAvailable(string name)
        {
            if (_names.Contains(name))
            {
                throw new ArgumentException($"A table named '{name}' already exists in this workbook.", nameof(name));
            }
        }

        internal int Claim(string name)
        {
            RequireAvailable(name);
            _names.Add(name);
            return _nextId++;
        }
    }
}
