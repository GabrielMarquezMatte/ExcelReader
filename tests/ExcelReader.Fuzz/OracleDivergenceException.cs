namespace ExcelReader.Fuzz
{
    /// <summary>
    /// Two decoders disagreed on the same input. Deliberately outside <see cref="FuzzOracle"/>'s
    /// sanctioned types, so <see cref="FuzzOracle.Guard"/> never swallows it.
    /// </summary>
    internal sealed class OracleDivergenceException : Exception
    {
        public OracleDivergenceException()
        {
        }

        public OracleDivergenceException(string message)
            : base(message)
        {
        }

        public OracleDivergenceException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
