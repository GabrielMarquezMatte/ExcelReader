#if NET9_0_OR_GREATER
namespace ExcelReader.Core.Parser
{
    // Immutable per-enumeration context handed to every IExcelRowModel<TSelf>.FromRow call. A normal
    // readonly struct (not ref struct), so it can be passed `in` freely (unlike Row itself).
    // Declared as a record struct purely for its compiler-synthesized equality members, satisfying
    // CA1815 without hand-written boilerplate; nothing in this codebase actually compares two
    // contexts. Kept non-positional (no primary constructor) so construction stays internal-only.
    public readonly record struct ExcelRowContext
    {
        public bool IsDate1904 { get; }
        public IFormatProvider FormatProvider { get; }

        internal ExcelRowContext(bool isDate1904, IFormatProvider formatProvider)
        {
            IsDate1904 = isDate1904;
            FormatProvider = formatProvider;
        }
    }
}
#endif
