namespace ExcelReader.Core.Parser.Internal
{
    // One resolved header->column->parser binding, shared by RowProjector<T> (class/struct models) and
    // NamedRefRowEnumerator<TModel,...> (ref struct models) via SparseRowProjection.
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
