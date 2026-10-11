namespace ExcelReader.Arrow
{
    /// <summary>Options for writing a record batch with <see cref="ArrowWriteExtensions"/>.</summary>
    public sealed class ArrowWriteOptions
    {
        /// <summary>Gets the default options: a sheet named <c>Sheet1</c> with a header row.</summary>
        public static ArrowWriteOptions Default { get; } = new();

        /// <summary>Gets the name of the sheet the batch is written to. Defaults to <c>Sheet1</c>.</summary>
        public string SheetName { get; init; } = "Sheet1";

        /// <summary>Gets whether the first row holds the batch's field names. Defaults to <see langword="true"/>.</summary>
        public bool WriteHeader { get; init; } = true;
    }
}
