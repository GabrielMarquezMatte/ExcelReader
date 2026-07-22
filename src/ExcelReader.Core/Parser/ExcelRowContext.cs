#if NET9_0_OR_GREATER
namespace ExcelReader.Core.Parser
{
    // Immutable per-enumeration context handed to every IExcelRowModel<TSelf>.FromRow call. A normal
    // readonly struct (not ref struct), so it can be passed `in` freely (unlike Row itself).
    public readonly struct ExcelRowContext : IEquatable<ExcelRowContext>
    {
        public bool IsDate1904 { get; }
        public IFormatProvider FormatProvider { get; }

        internal ExcelRowContext(bool isDate1904, IFormatProvider formatProvider)
        {
            IsDate1904 = isDate1904;
            FormatProvider = formatProvider;
        }

        public bool Equals(ExcelRowContext other)
        {
            return IsDate1904 == other.IsDate1904 && FormatProvider.Equals(other.FormatProvider);
        }

        public override bool Equals(object? obj)
        {
            return obj is ExcelRowContext other && Equals(other);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(IsDate1904, FormatProvider);
        }

        public static bool operator ==(ExcelRowContext left, ExcelRowContext right)
        {
            return left.Equals(right);
        }

        public static bool operator !=(ExcelRowContext left, ExcelRowContext right)
        {
            return !left.Equals(right);
        }
    }
}
#endif
