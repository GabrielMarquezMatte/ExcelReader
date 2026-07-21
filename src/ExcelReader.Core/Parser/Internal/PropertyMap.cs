namespace ExcelReader.Core.Parser.Internal
{
    internal readonly struct PropertyMap<T>
#if NET9_0_OR_GREATER
        where T : allows ref struct
#endif
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
        // The column must be present in the header.
        internal bool IsRequired { get; }
        // The cell must also be non-empty in every data row.
        internal bool RequireValue { get; }
    }
}
