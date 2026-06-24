namespace ExcelReader.Core.Parser.Internal
{
    internal readonly struct PropertyMap<T> where T : new()
    {
        internal PropertyMap(string[] names, ColumnParser<T> parser)
        {
            Names = names;
            Parser = parser;
        }

        internal string[] Names { get; }
        internal ColumnParser<T> Parser { get; }
    }
}
