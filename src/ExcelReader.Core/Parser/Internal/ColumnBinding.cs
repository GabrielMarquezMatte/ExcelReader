namespace ExcelReader.Core.Parser.Internal
{
    // One resolved header->column->parser binding, shared by RowProjector<T> (class/struct models) and
    // NamedRefRowEnumerator<TModel,...> (ref struct models, net9+) via SparseRowProjection. `allows ref
    // struct` only exists as a constraint kind on net9+; on net8 TModel is always a class/struct anyway
    // (NamedRefRowEnumerator itself is entirely #if NET9_0_OR_GREATER-gated), so the plain unconstrained
    // form there is exactly as capable.
    internal readonly struct ColumnBinding<TModel>
#if NET9_0_OR_GREATER
        where TModel : allows ref struct
#endif
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
