#if NET9_0_OR_GREATER
namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// Immutable per-enumeration context handed to every generated <c>IExcelRowModel&lt;TSelf&gt;.FromRow</c>
    /// call, carrying settings a row model needs to convert its own cells.
    /// </summary>
    /// <remarks>
    /// A normal <see langword="readonly"/> struct (not a ref struct), so it can be passed <c>in</c>
    /// freely, unlike <c>Row</c> itself. Declared as a record struct purely for its
    /// compiler-synthesized equality members (satisfying CA1815 without hand-written boilerplate; no
    /// code in this library actually compares two contexts), and kept non-positional (no primary
    /// constructor) so construction stays internal-only.
    /// </remarks>
    public readonly record struct ExcelRowContext
    {
        /// <summary>True when the source workbook uses the 1904 date system.</summary>
        public bool IsDate1904 { get; }
        /// <summary>The format provider to use when parsing cell values.</summary>
        public IFormatProvider FormatProvider { get; }

        internal ExcelRowContext(bool isDate1904, IFormatProvider formatProvider)
        {
            IsDate1904 = isDate1904;
            FormatProvider = formatProvider;
        }
    }
}
#endif
