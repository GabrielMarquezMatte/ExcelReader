namespace ExcelReader.NetFramework
{
    /// <summary>A native call failed.</summary>
    public sealed class ExcelReaderException : Exception
    {
        /// <summary>The native status code that caused the failure.</summary>
        public ExcelStatus Status { get; }

        /// <summary>Creates an exception with <see cref="ExcelStatus.Error"/> and a default message.</summary>
        public ExcelReaderException() : this(ExcelStatus.Error, "Native call failed.")
        {
        }

        /// <summary>Creates an exception with <see cref="ExcelStatus.Error"/>.</summary>
        /// <param name="message">The error message.</param>
        public ExcelReaderException(string message) : this(ExcelStatus.Error, message)
        {
        }

        /// <summary>Creates an exception with <see cref="ExcelStatus.Error"/> wrapping another exception.</summary>
        /// <param name="message">The error message.</param>
        /// <param name="innerException">The exception that caused this one.</param>
        public ExcelReaderException(string message, Exception innerException) : base(message, innerException)
        {
            Status = ExcelStatus.Error;
        }

        /// <summary>Creates an exception for a native status.</summary>
        /// <param name="status">The native status code.</param>
        /// <param name="message">The error message.</param>
        public ExcelReaderException(ExcelStatus status, string message) : base(message)
        {
            Status = status;
        }
    }
}
