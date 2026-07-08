namespace ExcelReader.Core.Writer
{
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

        public static XlsbCell Empty => default;

        public static XlsbCell Create(string? value)
        {
            return value is null
                ? Empty
                : new XlsbCell(XlsbCellKind.String, value, 0, boolean: false);
        }

        public static XlsbCell Create(bool value)
        {
            return new XlsbCell(XlsbCellKind.Boolean, text: null, number: 0, value);
        }

        public static XlsbCell Create(bool? value)
        {
            return value.HasValue ? Create(value.GetValueOrDefault()) : Empty;
        }

        public static XlsbCell Create(DateTime value)
        {
            return new XlsbCell(XlsbCellKind.Date, text: null, value.ToOADate(), boolean: false);
        }

        public static XlsbCell Create(DateTime? value)
        {
            return value.HasValue ? Create(value.GetValueOrDefault()) : Empty;
        }

        public static XlsbCell Create(int value)
        {
            return CreateNumber(value);
        }

        public static XlsbCell Create(int? value)
        {
            return value.HasValue ? CreateNumber(value.GetValueOrDefault()) : Empty;
        }

        public static XlsbCell Create(long value)
        {
            return CreateNumber(value);
        }

        public static XlsbCell Create(long? value)
        {
            return value.HasValue ? CreateNumber(value.GetValueOrDefault()) : Empty;
        }

        public static XlsbCell Create(float value)
        {
            return CreateNumber(value);
        }

        public static XlsbCell Create(float? value)
        {
            return value.HasValue ? CreateNumber(value.GetValueOrDefault()) : Empty;
        }

        public static XlsbCell Create(double value)
        {
            return CreateNumber(value);
        }

        public static XlsbCell Create(double? value)
        {
            return value.HasValue ? CreateNumber(value.GetValueOrDefault()) : Empty;
        }

        public static XlsbCell Create(decimal value)
        {
            return CreateNumber((double)value);
        }

        public static XlsbCell Create(decimal? value)
        {
            return value.HasValue ? CreateNumber((double)value.GetValueOrDefault()) : Empty;
        }

        public static XlsbCell Create<T>(T value)
            where T : IUtf8SpanFormattable
        {
            return CreateNumber(XlsbRowWriter.ToDouble(value));
        }

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
