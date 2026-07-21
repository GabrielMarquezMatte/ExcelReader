namespace ExcelReader.Core.Parser.Internal
{
    internal readonly struct HeaderMatch<T>
#if NET9_0_OR_GREATER
        where T : allows ref struct
#endif
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
