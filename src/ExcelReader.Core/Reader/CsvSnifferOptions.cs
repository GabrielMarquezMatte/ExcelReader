namespace ExcelReader.Core.Reader
{
    /// <summary>Options controlling how <see cref="CsvSniffer.Detect(ReadOnlySpan{byte}, CsvSnifferOptions)"/> infers a <see cref="CsvDialect"/>.</summary>
    public sealed record CsvSnifferOptions
    {
        // Defensive copy on every get: an array is mutable regardless of how the reference reached the
        // caller, and CsvSnifferOptions.Default is one process-lifetime singleton — a caller mutating
        // Default.CandidateDelimiters[0] in place (easy to do by accident, e.g. `options.CandidateDelimiters[0] = ...`
        // reading through Default) would corrupt every future default-options detection in the process.
        // The init accessor still stores the caller's array by reference, same cost as before for a
        // caller who builds their own instance and never touches Default.
        private readonly byte[] _candidateDelimiters = [(byte)',', (byte)';', (byte)'\t', (byte)'|'];
        private readonly byte[] _candidateQuotes = [(byte)'"', (byte)'\''];

        /// <summary>Gets the delimiters tried, in order. The order is the tie-break criterion when two candidates score equally. Defaults to <c>,</c> <c>;</c> tab <c>|</c>.</summary>
        public byte[] CandidateDelimiters
        {
            get => [.. _candidateDelimiters];
            init => _candidateDelimiters = value;
        }

        /// <summary>Gets the quote characters tried, in order. Defaults to <c>"</c> <c>'</c>.</summary>
        public byte[] CandidateQuotes
        {
            get => [.. _candidateQuotes];
            init => _candidateQuotes = value;
        }

        /// <summary>Gets the maximum number of lines from the sample considered. Defaults to 20.</summary>
        public int MaxSampleLines { get; init; } = 20;

        /// <summary>Gets the default options instance.</summary>
        public static CsvSnifferOptions Default { get; } = new();
    }
}
