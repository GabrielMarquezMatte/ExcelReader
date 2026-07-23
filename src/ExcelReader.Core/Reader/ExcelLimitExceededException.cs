namespace ExcelReader.Core.Reader
{
    /// <summary>
    /// Thrown when reading a workbook would exceed one of the resource limits configured on
    /// <see cref="ExcelReaderOptions"/> (e.g. decompressed size, a single cell's byte length, shared-string
    /// table size, or ZIP entry count), guarding against maliciously or accidentally oversized files.
    /// </summary>
    public sealed class ExcelLimitExceededException : Exception
    {
        /// <summary>Initializes a new instance of the <see cref="ExcelLimitExceededException"/> class.</summary>
        public ExcelLimitExceededException()
        {
        }

        /// <summary>Initializes a new instance of the <see cref="ExcelLimitExceededException"/> class with the given message.</summary>
        /// <param name="message">The message that describes the error.</param>
        public ExcelLimitExceededException(string message)
            : base(message)
        {
        }

        /// <summary>Initializes a new instance of the <see cref="ExcelLimitExceededException"/> class with the given message and inner exception.</summary>
        /// <param name="message">The message that describes the error.</param>
        /// <param name="innerException">The exception that caused this exception.</param>
        public ExcelLimitExceededException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        /// <summary>Initializes a new instance of the <see cref="ExcelLimitExceededException"/> class for a specific exceeded limit, building the message from the given values and populating <see cref="LimitName"/>, <see cref="Limit"/>, and <see cref="Actual"/>.</summary>
        /// <param name="limitName">The name of the exceeded limit (e.g. the corresponding <see cref="ExcelReaderOptions"/> property name).</param>
        /// <param name="limit">The configured limit, in bytes.</param>
        /// <param name="actual">The actual size, in bytes, that exceeded the limit.</param>
        public ExcelLimitExceededException(string limitName, long limit, long actual)
            : base($"{limitName} limit exceeded: {actual} bytes exceeds {limit} bytes.")
        {
            LimitName = limitName;
            Limit = limit;
            Actual = actual;
        }

        /// <summary>Gets the name of the limit that was exceeded.</summary>
        public string LimitName { get; } = string.Empty;

        /// <summary>Gets the configured limit, in bytes, that was exceeded.</summary>
        public long Limit { get; }

        /// <summary>Gets the actual size, in bytes, that exceeded <see cref="Limit"/>.</summary>
        public long Actual { get; }
    }
}
