namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// A workbook password. Deliberately a class rather than a record: a record's synthesized
    /// <c>ToString</c> would print the password, and this type is held by
    /// <see cref="ExcelReaderOptions"/>, which <i>is</i> a record — so one
    /// <c>logger.LogDebug("{Options}", options)</c> would leak it. <see cref="ToString"/> below is
    /// what makes that safe.
    /// </summary>
    /// <remarks>
    /// <b>On wiping secrets, plainly:</b> a password that reaches this type as a <see cref="string"/>
    /// cannot be wiped — .NET strings are immutable, movable by the GC, and possibly interned. This
    /// type does not pretend otherwise, and deliberately does not use <c>SecureString</c>, which is
    /// deprecated and only ever protected anything on Windows. What the library actually does is zero
    /// the UTF-16 buffer it allocates during key derivation and zero the derived key when the reader
    /// is disposed — that key, not the password, is what lives in memory for the reader's lifetime.
    /// </remarks>
    public sealed class ExcelPassword
    {
        private readonly char[] _chars;

        /// <summary>Initializes a new instance from a string.</summary>
        /// <param name="password">The password. Must not be <see langword="null"/>.</param>
        public ExcelPassword(string password)
        {
            ArgumentNullException.ThrowIfNull(password);
            _chars = password.ToCharArray();
        }

        /// <summary>Initializes a new instance from a character span, for callers avoiding a string.</summary>
        /// <param name="password">The password characters, copied into this instance.</param>
        public ExcelPassword(ReadOnlySpan<char> password)
        {
            _chars = password.ToArray();
        }

        /// <summary>Converts a string to an <see cref="ExcelPassword"/>.</summary>
        /// <param name="password">The password.</param>
        public static implicit operator ExcelPassword(string password)
        {
            return new(password);
        }

        /// <summary>Creates an <see cref="ExcelPassword"/> from a string. Named alternative to the implicit conversion.</summary>
        /// <param name="password">The password.</param>
        public static ExcelPassword FromString(string password)
        {
            return new(password);
        }

        /// <summary>Returns a redacted placeholder, never the password itself.</summary>
        public override string ToString()
        {
            return "ExcelPassword(set)";
        }

        internal ReadOnlySpan<char> Chars => _chars;
    }
}
