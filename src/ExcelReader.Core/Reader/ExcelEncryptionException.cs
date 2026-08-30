using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Reader
{
    /// <summary>Why an encrypted workbook could not be opened.</summary>
    public enum ExcelEncryptionReason
    {
        /// <summary>The workbook is encrypted and no password was supplied.</summary>
        PasswordRequired,
        /// <summary>The supplied password did not match the workbook's verifier.</summary>
        PasswordIncorrect,
        /// <summary>The workbook uses an encryption scheme this library does not support.</summary>
        UnsupportedScheme,
        /// <summary>The workbook's integrity check failed: it is corrupt or has been tampered with.</summary>
        IntegrityFailure,
    }

    /// <summary>
    /// Thrown when an encrypted workbook cannot be opened. Inspect <see cref="Reason"/> to decide what
    /// to do: only <see cref="ExcelEncryptionReason.PasswordRequired"/> and
    /// <see cref="ExcelEncryptionReason.PasswordIncorrect"/> are worth prompting a user about — the
    /// other two are terminal, and retrying with another password will never help.
    /// </summary>
    public sealed class ExcelEncryptionException : IOException
    {
        // CA1032/RCS1194 require the three standard exception constructors alongside the domain
        // ones below, matching the convention already used by ExcelParseException. None of these are
        // ever constructed by this library — Reason would be left at its default (PasswordRequired) —
        // but they exist so this remains a well-behaved Exception type for callers/serializers.
        /// <summary>Creates an exception with no message.</summary>
        [ExcludeFromCodeCoverage]
        public ExcelEncryptionException()
        {
        }

        /// <summary>Creates an exception with the given message.</summary>
        [ExcludeFromCodeCoverage]
        public ExcelEncryptionException(string message) : base(message)
        {
        }

        /// <summary>Creates an exception with the given message and inner exception.</summary>
        [ExcludeFromCodeCoverage]
        public ExcelEncryptionException(string message, Exception innerException) : base(message, innerException)
        {
        }

        /// <summary>Initializes a new instance with the given reason and message.</summary>
        /// <param name="reason">Why the workbook could not be opened.</param>
        /// <param name="message">A human-readable description.</param>
        public ExcelEncryptionException(ExcelEncryptionReason reason, string message) : base(message)
        {
            Reason = reason;
        }

        /// <summary>Initializes a new instance with the given reason, message and inner exception.</summary>
        /// <param name="reason">Why the workbook could not be opened.</param>
        /// <param name="message">A human-readable description.</param>
        /// <param name="innerException">The underlying failure.</param>
        public ExcelEncryptionException(ExcelEncryptionReason reason, string message, Exception innerException)
            : base(message, innerException)
        {
            Reason = reason;
        }

        /// <summary>
        /// Initializes a new instance with the given message and HRESULT. This is used when the underlying COM call fails, and the HRESULT is preserved for diagnostic purposes.
        /// </summary>
        /// <param name="message">A human-readable description.</param>
        /// <param name="hresult">The HRESULT returned by the underlying COM call.</param>
        [ExcludeFromCodeCoverage]
        public ExcelEncryptionException(string? message, int hresult) : base(message, hresult)
        {
        }

        /// <summary>Gets the reason the workbook could not be opened.</summary>
        public ExcelEncryptionReason Reason { get; }
    }
}
