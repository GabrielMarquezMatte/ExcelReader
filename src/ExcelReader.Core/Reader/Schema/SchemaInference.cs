using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Reader.Schema
{
    /// <summary>
    /// Guesses a column schema by sampling a sheet's rows. Backs
    /// <see cref="Excel.InferSchema(IExcelRowReader, int, int)"/> and the native <c>xl_infer_schema</c>
    /// export, so both answer identically by construction.
    /// </summary>
    internal static class SchemaInference
    {
        /// <summary>
        /// Samples <paramref name="rows"/> from its current position and guesses one
        /// <see cref="ExcelColumnSchema"/> per column. Every guess comes from the sampled cells' own
        /// <see cref="CellType"/> tag, and, with <paramref name="parseText"/>, from the exact shape of text cells.
        /// </summary>
        /// <param name="rows">The sheet's rows, positioned at the first row to sample (the header row if any, or the first data row).</param>
        /// <param name="isDate1904">Whether the sheet uses the 1904 date system, which shifts all date/time values by 1462 days.</param>
        /// <param name="headerRow">1-based row number to take column names from; 0 means "no header",
        /// so every returned schema is index-only.</param>
        /// <param name="sampleSize">How many rows after the header to inspect. Must be positive.</param>
        /// <param name="parseText">Whether a text cell counts as an integer, decimal, boolean, ISO date or ISO date-time when its text has exactly that shape.</param>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="headerRow"/> is negative, or
        /// <paramref name="sampleSize"/> is not positive.</exception>
        /// <exception cref="ArgumentException">The sheet has fewer rows than <paramref name="headerRow"/>.</exception>
        internal static ExcelColumnSchema[] Infer(IExcelRowEnumerator rows, bool isDate1904, int headerRow, int sampleSize, bool parseText = false)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(headerRow);
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(sampleSize);
            List<string?> names = [];
            List<ColumnStat> stats = [];
            if (headerRow > 0 && !TryReadHeader(rows, headerRow, names, stats, out string? headerError))
            {
                throw new ArgumentException(headerError, nameof(headerRow));
            }
            int dataRowCount = SampleDataRows(rows, sampleSize, isDate1904, parseText, names, stats);
            MarkSparseColumnsNullable(CollectionsMarshal.AsSpan(stats), dataRowCount);
            ExcelColumnSchema[] schema = new ExcelColumnSchema[names.Count];
            for (int i = 0; i < schema.Length; i++)
            {
                schema[i] = new ExcelColumnSchema
                {
                    Index = i,
                    Name = names[i],
                    Type = stats[i].InferType(parseText),
                    IsNullable = stats[i].SawEmpty,
                };
            }
            return schema;
        }

        internal static bool TrySkipToHeaderRow(IExcelRowEnumerator rows, int headerRow, [NotNullWhen(false)] out string? error)
        {
            error = null;
            for (int rowNumber = 1; rowNumber <= headerRow; rowNumber++)
            {
                if (!rows.MoveNext())
                {
                    error = $"sheet has fewer than {headerRow} row(s); cannot resolve header_row.";
                    return false;
                }
            }
            return true;
        }

        private static void EnsureCapacity(List<string?> names, List<ColumnStat> stats, int index)
        {
            while (names.Count <= index)
            {
                names.Add(null);
                stats.Add(default);
            }
        }

        private static bool TryReadHeader(IExcelRowEnumerator rows, int headerRow, List<string?> names, List<ColumnStat> stats, out string? error)
        {
            if (!TrySkipToHeaderRow(rows, headerRow, out error))
            {
                return false;
            }
            Row header = rows.Current;
            foreach (RowCell cell in header.Cells)
            {
                EnsureCapacity(names, stats, cell.ColumnIndex);
                string text = cell.Value.GetString().Trim();
                names[cell.ColumnIndex] = text.Length == 0 ? null : text;
            }
            return true;
        }

        private static int SampleDataRows(IExcelRowEnumerator rows, int sampleSize, bool isDate1904, bool parseText, List<string?> names, List<ColumnStat> stats)
        {
            int dataRowCount = 0;
            for (int i = 0; i < sampleSize && rows.MoveNext(); i++)
            {
                dataRowCount++;
                Row row = rows.Current;
                foreach (RowCell cell in row.Cells)
                {
                    EnsureCapacity(names, stats, cell.ColumnIndex);
                    Cell value = cell.Value;
                    CollectionsMarshal.AsSpan(stats)[cell.ColumnIndex].Observe(in value, isDate1904, parseText);
                }
            }
            return dataRowCount;
        }

        private static void MarkSparseColumnsNullable(Span<ColumnStat> stats, int dataRowCount)
        {
            foreach (ref ColumnStat stat in stats)
            {
                if (stat.SeenCount < dataRowCount)
                {
                    stat.SawEmpty = true;
                }
            }
        }

        [StructLayout(LayoutKind.Auto)]
        private struct ColumnStat
        {
            internal int SeenCount;
            internal bool SawEmpty;
            internal bool SawString;
            internal bool SawNumber;
            internal bool SawDate;
            internal bool SawTimestamp;
            internal bool SawBool;
            internal bool SawFormulaOrError;
            internal bool SawNonIntegralNumber;
            internal bool SawNonBinaryNumber;

            internal void Observe(in Cell cell, bool isDate1904, bool parseText)
            {
                SeenCount++;
                switch (cell.Type)
                {
                    case CellType.Empty:
                        SawEmpty = true;
                        break;
                    case CellType.ExcelString:
                        if (!parseText || !ObserveText(in cell, isDate1904))
                        {
                            SawString = true;
                        }
                        break;
                    case CellType.Number:
                        SawNumber = true;
                        if (!ExcelCellReaders.Parsable<long>(in cell, isDate1904, CultureInfo.InvariantCulture, out _))
                        {
                            SawNonIntegralNumber = true;
                        }
                        SawNonBinaryNumber |= !IsBinaryDigit(cell.Value);
                        break;
                    case CellType.Date:
                        SawDate = true;
                        break;
                    case CellType.Boolean:
                        SawBool = true;
                        break;
                    default:
                        SawFormulaOrError = true;
                        break;
                }
            }

            // Stricter than the typed converters on purpose: their fallbacks (DateTime.TryParse,
            // long.TryParse with whitespace) would type ID and ratio columns as dates or numbers.
            private bool ObserveText(in Cell cell, bool isDate1904)
            {
                ReadOnlySpan<byte> text = cell.Value;
                bool canonical = HasCanonicalIntegerPart(text, out bool allDigits);
                if (canonical && allDigits && ExcelCellReaders.Parsable<long>(in cell, isDate1904, CultureInfo.InvariantCulture, out _))
                {
                    SawNumber = true;
                    SawNonBinaryNumber |= !IsBinaryDigit(text);
                    return true;
                }
                if (canonical && FastDouble.TryParse(text, out _))
                {
                    SawNumber = true;
                    SawNonIntegralNumber = true;
                    SawNonBinaryNumber = true;
                    return true;
                }
                if (Ascii.EqualsIgnoreCase(text, "true"u8) || Ascii.EqualsIgnoreCase(text, "false"u8))
                {
                    SawBool = true;
                    return true;
                }
                if (!FastDate.TryParse(text, out _))
                {
                    return false;
                }
                if (text.Length == 10)
                {
                    SawDate = true;
                    return true;
                }
                SawTimestamp = true;
                return true;
            }

            private static bool HasCanonicalIntegerPart(ReadOnlySpan<byte> text, out bool allDigits)
            {
                int start = !text.IsEmpty && text[0] is (byte)'-' or (byte)'+' ? 1 : 0;
                int end = start;
                while (end < text.Length && (uint)(text[end] - (byte)'0') <= 9)
                {
                    end++;
                }
                allDigits = end > start && end == text.Length;
                return end - start <= 1 || text[start] != (byte)'0';
            }

            private static bool IsBinaryDigit(ReadOnlySpan<byte> text)
            {
                return text.Length == 1 && text[0] is (byte)'0' or (byte)'1';
            }

            internal readonly ExcelColumnType InferType(bool parseText)
            {
                bool boolWithBinaryNumbers = parseText && SawBool && SawNumber && !SawNonIntegralNumber && !SawNonBinaryNumber;
                int kinds = (SawString ? 1 : 0) + (SawNumber && !boolWithBinaryNumbers ? 1 : 0)
                    + (SawDate || SawTimestamp ? 1 : 0) + (SawBool ? 1 : 0);
                if (SawFormulaOrError || kinds != 1 || SawString)
                {
                    return ExcelColumnType.StringColumn;
                }
                if (SawTimestamp)
                {
                    return ExcelColumnType.TimestampColumn;
                }
                if (SawDate)
                {
                    return ExcelColumnType.DateColumn;
                }
                if (SawBool)
                {
                    return ExcelColumnType.BoolColumn;
                }
                return SawNonIntegralNumber ? ExcelColumnType.Float64Column : ExcelColumnType.Int64Column;
            }
        }
    }
}
