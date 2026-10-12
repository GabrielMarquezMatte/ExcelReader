namespace ExcelReader.Core.Reader.Schema
{
    /// <summary>Options for guessing a column schema with <see cref="Excel.InferSchema"/>.</summary>
    public sealed class SchemaInferenceOptions
    {
        /// <summary>Gets the default options: header on row 1, 100 sampled rows, text left as text.</summary>
        public static SchemaInferenceOptions Default { get; } = new();

        /// <summary>
        /// Gets the 1-based row number to take column names from. <c>0</c> means "no header", so every
        /// returned schema is addressable only by <see cref="ExcelColumnSchema.Index"/>. Defaults to <c>1</c>.
        /// </summary>
        public int HeaderRow { get; init; } = 1;

        /// <summary>Gets how many rows after the header to inspect. Defaults to <c>100</c>.</summary>
        public int SampleSize { get; init; } = 100;

        /// <summary>
        /// Gets whether a text cell (every CSV field, or a number stored as text) counts as an integer, a
        /// decimal, <c>true</c>/<c>false</c>, or an ISO-8601 date or date-time when its text has exactly that
        /// shape. Numbers with a leading zero (<c>00123</c>), scientific notation (<c>12E4</c>), padded or
        /// culture-formatted numbers, and non-ISO dates stay text. A column only gets a non-text type when
        /// every sampled value converts to it. Defaults to <see langword="false"/>.
        /// </summary>
        public bool ParseText { get; init; }
    }
}
