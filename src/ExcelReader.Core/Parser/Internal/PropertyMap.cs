namespace ExcelReader.Core.Parser.Internal
{
    internal readonly struct PropertyMap<T>
        where T : allows ref struct
    {
        internal PropertyMap(string[] names, ColumnParser<T> parser, bool isRequired, bool requireValue)
        {
            Names = names;
            Parser = parser;
            IsRequired = isRequired;
            RequireValue = requireValue;
        }

        internal string[] Names { get; }
        internal ColumnParser<T> Parser { get; }
        internal bool IsRequired { get; }
        internal bool RequireValue { get; }
    }
}
