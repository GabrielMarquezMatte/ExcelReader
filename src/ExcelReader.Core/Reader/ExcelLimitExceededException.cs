namespace ExcelReader.Core.Reader
{
    public sealed class ExcelLimitExceededException : Exception
    {
        public ExcelLimitExceededException()
        {
        }

        public ExcelLimitExceededException(string message)
            : base(message)
        {
        }

        public ExcelLimitExceededException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        public ExcelLimitExceededException(string limitName, long limit, long actual)
            : base($"{limitName} limit exceeded: {actual} bytes exceeds {limit} bytes.")
        {
            LimitName = limitName;
            Limit = limit;
            Actual = actual;
        }

        public string LimitName { get; } = string.Empty;
        public long Limit { get; }
        public long Actual { get; }
    }
}
