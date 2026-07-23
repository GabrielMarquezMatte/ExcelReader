namespace ExcelReader.Core.Writer
{
    /// <summary>
    /// An immutable, typed cell value for <see cref="XlsbSheetWriter.WriteRow(ReadOnlySpan{XlsbCell})"/>,
    /// created via one of the <see cref="Create(string?)"/> factory overloads.
    /// </summary>
    public readonly record struct XlsbCell
    {
        internal XlsbCell(XlsbCellKind kind, string? text, double number, bool boolean)
        {
            Kind = kind;
            Text = text;
            Number = number;
            Boolean = boolean;
        }

        internal XlsbCellKind Kind { get; }
        internal string? Text { get; }
        internal double Number { get; }
        internal bool Boolean { get; }

        /// <summary>An empty cell, written as a gap in the row.</summary>
        public static XlsbCell Empty => default;

        /// <summary>Creates a string cell from <paramref name="value"/>, or <see cref="Empty"/> if it is <see langword="null"/>.</summary>
        public static XlsbCell Create(string? value)
        {
            return value is null
                ? Empty
                : new XlsbCell(XlsbCellKind.String, value, 0, boolean: false);
        }

        /// <summary>Creates a boolean cell from <paramref name="value"/>.</summary>
        public static XlsbCell Create(bool value)
        {
            return new XlsbCell(XlsbCellKind.Boolean, text: null, number: 0, value);
        }

        /// <summary>Creates a boolean cell from <paramref name="value"/>, or <see cref="Empty"/> if it has no value.</summary>
        public static XlsbCell Create(bool? value)
        {
            return value.HasValue ? Create(value.GetValueOrDefault()) : Empty;
        }

        /// <summary>Creates a date cell from <paramref name="value"/>.</summary>
        public static XlsbCell Create(DateTime value)
        {
            return new XlsbCell(XlsbCellKind.Date, text: null, value.ToOADate(), boolean: false);
        }

        /// <summary>Creates a date cell from <paramref name="value"/>, or <see cref="Empty"/> if it has no value.</summary>
        public static XlsbCell Create(DateTime? value)
        {
            return value.HasValue ? Create(value.GetValueOrDefault()) : Empty;
        }

        /// <summary>Creates a numeric cell from <paramref name="value"/>.</summary>
        public static XlsbCell Create(int value)
        {
            return CreateNumber(value);
        }

        /// <summary>Creates a numeric cell from <paramref name="value"/>, or <see cref="Empty"/> if it has no value.</summary>
        public static XlsbCell Create(int? value)
        {
            return value.HasValue ? CreateNumber(value.GetValueOrDefault()) : Empty;
        }

        /// <summary>Creates a numeric cell from <paramref name="value"/>.</summary>
        public static XlsbCell Create(long value)
        {
            return CreateNumber(value);
        }

        /// <summary>Creates a numeric cell from <paramref name="value"/>, or <see cref="Empty"/> if it has no value.</summary>
        public static XlsbCell Create(long? value)
        {
            return value.HasValue ? CreateNumber(value.GetValueOrDefault()) : Empty;
        }

        /// <summary>Creates a numeric cell from <paramref name="value"/>.</summary>
        public static XlsbCell Create(float value)
        {
            return CreateNumber(value);
        }

        /// <summary>Creates a numeric cell from <paramref name="value"/>, or <see cref="Empty"/> if it has no value.</summary>
        public static XlsbCell Create(float? value)
        {
            return value.HasValue ? CreateNumber(value.GetValueOrDefault()) : Empty;
        }

        /// <summary>Creates a numeric cell from <paramref name="value"/>.</summary>
        public static XlsbCell Create(double value)
        {
            return CreateNumber(value);
        }

        /// <summary>Creates a numeric cell from <paramref name="value"/>, or <see cref="Empty"/> if it has no value.</summary>
        public static XlsbCell Create(double? value)
        {
            return value.HasValue ? CreateNumber(value.GetValueOrDefault()) : Empty;
        }

        /// <summary>Creates a numeric cell from <paramref name="value"/>.</summary>
        public static XlsbCell Create(decimal value)
        {
            return CreateNumber((double)value);
        }

        /// <summary>Creates a numeric cell from <paramref name="value"/>, or <see cref="Empty"/> if it has no value.</summary>
        public static XlsbCell Create(decimal? value)
        {
            return value.HasValue ? CreateNumber((double)value.GetValueOrDefault()) : Empty;
        }

        /// <summary>Creates a numeric cell from <paramref name="value"/>, formatting it via <typeparamref name="T"/>'s UTF-8 span formatting.</summary>
        public static XlsbCell Create<T>(T value)
            where T : IUtf8SpanFormattable
        {
            return CreateNumber(XlsbRowWriter.ToDouble(value));
        }

        /// <summary>Creates a numeric cell from <paramref name="value"/>, or <see cref="Empty"/> if it has no value.</summary>
        public static XlsbCell Create<T>(T? value)
            where T : struct, IUtf8SpanFormattable
        {
            return value.HasValue ? Create(value.GetValueOrDefault()) : Empty;
        }

        private static XlsbCell CreateNumber(double value)
        {
            return new XlsbCell(XlsbCellKind.Number, text: null, value, boolean: false);
        }
    }

    internal enum XlsbCellKind
    {
        Empty,
        String,
        Boolean,
        Number,
        Date,
    }
}
