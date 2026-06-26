namespace ExcelReader.Core.Parser.Internal
{
    internal readonly struct HeaderMatch<T>
    {
        internal HeaderMatch(int propertyIndex, int aliasIndex, ColumnParser<T> parser)
        {
            PropertyIndex = propertyIndex;
            AliasIndex = aliasIndex;
            Parser = parser;
        }

        internal int PropertyIndex { get; }
        internal int AliasIndex { get; }
        internal ColumnParser<T> Parser { get; }
    }
}
