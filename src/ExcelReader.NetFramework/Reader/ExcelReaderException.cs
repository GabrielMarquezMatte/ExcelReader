using System.Runtime.Serialization;
using ExcelReader.NetFramework.Parser.Internal;

namespace ExcelReader.NetFramework.Reader
{
    /// <summary>A native call failed.</summary>
    [Serializable]
    public sealed class ExcelReaderException : Exception
    {
        private const string StatusKey = "Status";

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

        private ExcelReaderException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
            Status = (ExcelStatus)info.GetInt32(StatusKey);
        }

        /// <inheritdoc />
        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            base.GetObjectData(info, context);
            info.AddValue(StatusKey, (int)Status);
        }
    }
}
