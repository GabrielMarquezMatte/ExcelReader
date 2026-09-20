namespace ExcelReader.Core.Parser.Internal
{
    internal readonly struct ColumnBinding<TModel>
        where TModel : allows ref struct
    {
        internal ColumnBinding(int column, ColumnParser<TModel> parser, bool requireValue, string name)
        {
            Column = column;
            Parser = parser;
            RequireValue = requireValue;
            Name = name;
        }

        internal int Column { get; }
        internal ColumnParser<TModel> Parser { get; }
        internal bool RequireValue { get; }
        internal string Name { get; }
    }
}
