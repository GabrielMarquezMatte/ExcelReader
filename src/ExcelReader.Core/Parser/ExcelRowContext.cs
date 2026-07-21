#if NET9_0_OR_GREATER
namespace ExcelReader.Core.Parser
{
    // Immutable per-enumeration context handed to every IExcelRowModel<TSelf>.FromRow call. A normal
    // readonly struct (not ref struct), so it can be passed `in` freely (unlike Row itself).
    public readonly struct ExcelRowContext
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
