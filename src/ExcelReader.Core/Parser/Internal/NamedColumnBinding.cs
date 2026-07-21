#if NET9_0_OR_GREATER
namespace ExcelReader.Core.Parser.Internal
{
    // One resolved header->column->parser binding for NamedRefRowEnumerator. Mirrors RowProjector<T>'s
    // private ColumnBinding<TModel>, kept as its own top-level type since RowProjector<T> can't be
    // reused directly for a ref-struct TModel (see NamedRefRowEnumerator's remarks).
    internal readonly struct NamedColumnBinding<T>
        where T : allows ref struct
    {
        internal NamedColumnBinding(int column, ColumnParser<T> parser, bool requireValue, string name)
        {
            Column = column;
            Parser = parser;
            RequireValue = requireValue;
            Name = name;
        }

        internal int Column { get; }
        internal ColumnParser<T> Parser { get; }
        internal bool RequireValue { get; }
        internal string Name { get; }
    }
}
#endif
