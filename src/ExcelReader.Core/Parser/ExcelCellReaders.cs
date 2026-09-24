using System.Diagnostics.CodeAnalysis;
using ExcelReader.Core.Parser.Internal;
using ExcelReader.Core.Reader;

namespace ExcelReader.Core.Parser
{
    /// <summary>
    /// The built-in cell readers <c>ExcelParser&lt;T&gt;</c>'s reflection path already uses, exposed as
    /// <see cref="ExcelCellReader{TValue}"/> instances for a source-generated or hand-written
    /// <see cref="ExcelRowMapBuilder{T}"/> map to plug into <see cref="ExcelRowMapBuilder{T}.Property{TValue}"/>
    /// without duplicating the conversion logic.
    /// </summary>
    public static class ExcelCellReaders
    {
        /// <summary>Reads a cell as text.</summary>
        [SuppressMessage("Naming", "CA1720:Identifier contains type name",
            Justification = "Matches the property type it reads, mirroring the other members of this class (Bool, DateTimeSerial, ...).")]
        public static readonly ExcelCellReader<string> String = ReadString;

        /// <summary>Reads a cell as <c>"1"</c>/<c>"0"</c> or <c>"true"</c>/<c>"false"</c> (case-insensitive).</summary>
        public static readonly ExcelCellReader<bool> Bool = ColumnParserFactory.ReadBool;

        /// <summary>Reads a cell as an Excel date/time serial number.</summary>
        public static readonly ExcelCellReader<DateTime> DateTimeSerial = ColumnParserFactory.ReadDateTime;

        /// <summary>Reads a cell as date/time text (ISO or culture format) rather than a serial number — the only option for CSV, which has no serial date form.</summary>
        public static readonly ExcelCellReader<DateTime> DateTimeText = ColumnParserFactory.ReadTextDateTime;

        /// <summary>
        /// Reads a cell as an Excel date/time serial number, falling back to date/time text when the
        /// cell isn't numeric. Used where a single map has to work for both a serial-number source
        /// (XLSX/XLSB/XLS) and CSV (which has no serial date form and writes ISO text instead) — see
        /// <see cref="ExcelParser.Generated{T}"/>, which builds one map for every
        /// reader. The cost: a CSV cell that is only digits (e.g. an Excel serial typed as plain text)
        /// is read as a serial number, not as date text — there is no way to tell those two apart from
        /// the cell alone.
        /// </summary>
        public static readonly ExcelCellReader<DateTime> DateTimeAuto = ReadDateTimeAuto;

        /// <summary>Reads a cell as an Excel date serial number, truncated to its date component.</summary>
        public static readonly ExcelCellReader<DateOnly> DateOnlySerial = ColumnParserFactory.ReadDateOnly;

        /// <summary>Reads a cell as date-only text rather than a serial number — the only option for CSV.</summary>
        public static readonly ExcelCellReader<DateOnly> DateOnlyText = ColumnParserFactory.ReadTextDateOnly;

        /// <summary>Reads a cell as an Excel date serial number, falling back to date-only text when the cell isn't numeric — the <see cref="DateOnly"/> counterpart of <see cref="DateTimeAuto"/>, same trade-off.</summary>
        public static readonly ExcelCellReader<DateOnly> DateOnlyAuto = ReadDateOnlyAuto;

        /// <summary>Reads a cell as an Excel time-of-day serial number (the fractional part of a day).</summary>
        public static readonly ExcelCellReader<TimeOnly> TimeOnlySerial = ColumnParserFactory.ReadTimeOnly;

        /// <summary>Reads a cell as time-only text rather than a serial number — the only option for CSV.</summary>
        public static readonly ExcelCellReader<TimeOnly> TimeOnlyText = ColumnParserFactory.ReadTextTimeOnly;

        /// <summary>Reads a cell as an Excel time-of-day serial number, falling back to time-only text when the cell isn't numeric — the <see cref="TimeOnly"/> counterpart of <see cref="DateTimeAuto"/>, same trade-off.</summary>
        public static readonly ExcelCellReader<TimeOnly> TimeOnlyAuto = ReadTimeOnlyAuto;


        private static bool ReadString(in Cell cell, bool isDate1904, IFormatProvider provider, out string value)
        {
            value = cell.GetString();
            return true;
        }

        private static bool ReadDateTimeAuto(in Cell cell, bool isDate1904, IFormatProvider provider, out DateTime value)
        {
            return ColumnParserFactory.ReadDateTime(in cell, isDate1904, provider, out value)
                || ColumnParserFactory.ReadTextDateTime(in cell, isDate1904, provider, out value);
        }

        private static bool ReadDateOnlyAuto(in Cell cell, bool isDate1904, IFormatProvider provider, out DateOnly value)
        {
            return ColumnParserFactory.ReadDateOnly(in cell, isDate1904, provider, out value)
                || ColumnParserFactory.ReadTextDateOnly(in cell, isDate1904, provider, out value);
        }

        private static bool ReadTimeOnlyAuto(in Cell cell, bool isDate1904, IFormatProvider provider, out TimeOnly value)
        {
            return ColumnParserFactory.ReadTimeOnly(in cell, isDate1904, provider, out value)
                || ColumnParserFactory.ReadTextTimeOnly(in cell, isDate1904, provider, out value);
        }

        /// <summary>
        /// Reads a cell as any type <see cref="Cell.TryParse{T}"/> supports: every integral type,
        /// <see cref="float"/>/<see cref="double"/>/<see cref="decimal"/>, and
        /// <see cref="Guid"/>.
        /// </summary>
        /// <typeparam name="TValue">The value type to parse.</typeparam>
        public static bool Parsable<TValue>(in Cell cell, bool isDate1904, IFormatProvider provider, [MaybeNullWhen(false)] out TValue value)
            where TValue : IUtf8SpanParsable<TValue>
        {
            return cell.TryParse(provider, out value);
        }

        /// <summary>
        /// Reads a cell by decoding its text and calling <see cref="ISpanParsable{TSelf}.TryParse(ReadOnlySpan{char}, IFormatProvider?, out TSelf)"/>,
        /// for types such as <see cref="TimeSpan"/> and <see cref="DateTimeOffset"/> that have no UTF-8 parser.
        /// </summary>
        /// <typeparam name="TValue">The value type to parse.</typeparam>
        public static bool SpanParsable<TValue>(in Cell cell, bool isDate1904, IFormatProvider provider, [MaybeNullWhen(false)] out TValue value)
            where TValue : ISpanParsable<TValue>
        {
            return ColumnParserFactory.TryParseSpanParsable(in cell, provider, out value);
        }

        /// <summary>Reads a cell's raw UTF-8 bytes without allocating. The span aliases the reader's buffer and is valid only until the reader advances.</summary>
        public static bool Utf8(scoped in Cell cell, bool isDate1904, IFormatProvider provider, out ReadOnlySpan<byte> value)
        {
            value = cell.Value;
            return true;
        }

        /// <summary>Reads a cell as an enum, by its declared value name (case-insensitive) or its underlying numeric value.</summary>
        /// <typeparam name="TEnum">The enum type to parse.</typeparam>
        public static bool Enum<TEnum>(in Cell cell, bool isDate1904, IFormatProvider provider, out TEnum value)
            where TEnum : struct, Enum
        {
            return ColumnParserFactory.TryParseEnum(in cell, out value);
        }
    }
}
