using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Parser;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Core.Reader
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
        /// <see cref="CellType"/> tag — no text sniffing.
        /// </summary>
        /// <param name="rows">The sheet's rows, positioned at the first row to sample (the header row if any, or the first data row).</param>
        /// <param name="isDate1904">Whether the sheet uses the 1904 date system, which shifts all date/time values by 1462 days.</param>
        /// <param name="headerRow">1-based row number to take column names from; 0 means "no header",
        /// so every returned schema is index-only.</param>
        /// <param name="sampleSize">How many rows after the header to inspect. Must be positive.</param>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="headerRow"/> is negative, or
        /// <paramref name="sampleSize"/> is not positive.</exception>
        /// <exception cref="ArgumentException">The sheet has fewer rows than <paramref name="headerRow"/>.</exception>
        internal static ExcelColumnSchema[] Infer(IExcelRowEnumerator rows, bool isDate1904, int headerRow, int sampleSize)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(headerRow);
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(sampleSize);
            List<string?> names = [];
            List<ColumnStat> stats = [];
            if (headerRow > 0 && !TryReadHeader(rows, headerRow, names, stats, out string? headerError))
            {
                throw new ArgumentException(headerError, nameof(headerRow));
            }
            int dataRowCount = SampleDataRows(rows, sampleSize, isDate1904, names, stats);
            MarkSparseColumnsNullable(CollectionsMarshal.AsSpan(stats), dataRowCount);
            ExcelColumnSchema[] schema = new ExcelColumnSchema[names.Count];
            for (int i = 0; i < schema.Length; i++)
            {
                schema[i] = new ExcelColumnSchema
                {
                    Index = i,
                    Name = names[i],
                    Type = stats[i].InferType(),
                    IsNullable = stats[i].SawEmpty,
                };
            }
            return schema;
        }

        // Advances `rows` so that `rows.Current` is the header row itself. Shared with the typed
        // parser in ExcelReader.Native so the two cannot drift on the row arithmetic or on the
        // message a too-short sheet produces.
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

        // Advances `rows` past headerRow and records each populated header cell's trimmed text as that
        // column's name. A blank header cell leaves its column name null (index-based) rather than an
        // empty string — an empty name would fail xl_parse_typed's own "blank name" validation later.
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

        // Reads up to sampleSize rows after the header, folding each populated cell's type into that
        // column's ColumnStat. Returns how many data rows were actually sampled (fewer than sampleSize
        // at end of sheet), used by MarkSparseColumnsNullable below.
        private static int SampleDataRows(IExcelRowEnumerator rows, int sampleSize, bool isDate1904, List<string?> names, List<ColumnStat> stats)
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
                    CollectionsMarshal.AsSpan(stats)[cell.ColumnIndex].Observe(in value, isDate1904);
                }
            }
            return dataRowCount;
        }

        // A column absent from some sampled rows never triggers ColumnStat.Observe for them — which is
        // exactly what row[index] would have reported for it: CellType.Empty. Comparing each column's
        // Observe count against the number of rows actually sampled catches that without a second,
        // O(columns) walk of every row via the indexer.
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

        // Accumulates what kinds of CellType a column's sampled cells held, enough to guess a
        // xl_parse_typed ColumnSpec without ever storing a cell's value past its own row.
        [StructLayout(LayoutKind.Auto)]
        private struct ColumnStat
        {
            internal int SeenCount;
            internal bool SawEmpty;
            internal bool SawString;
            internal bool SawNumber;
            internal bool SawDate;
            internal bool SawBool;
            internal bool SawFormulaOrError;
            internal bool SawNonIntegralNumber;

            internal void Observe(in Cell cell, bool isDate1904)
            {
                SeenCount++;
                switch (cell.Type)
                {
                    case CellType.Empty:
                        SawEmpty = true;
                        break;
                    case CellType.ExcelString:
                        SawString = true;
                        break;
                    case CellType.Number:
                        SawNumber = true;
                        if (!ExcelCellReaders.Parsable<long>(in cell, isDate1904, CultureInfo.InvariantCulture, out _))
                        {
                            SawNonIntegralNumber = true;
                        }
                        break;
                    case CellType.Date:
                        SawDate = true;
                        break;
                    case CellType.Boolean:
                        SawBool = true;
                        break;
                    default: // Formula, Error — the cached result was never sampled as a plain value.
                        SawFormulaOrError = true;
                        break;
                }
            }

            // A column only gets a non-string guess when every sampled cell agreed on one single kind.
            // A real mix, any formula/error result, or nothing seen at all falls back to the string
            // type, since it is the only one able to represent every one of those verbatim.
            internal readonly ExcelColumnType InferType()
            {
                int kinds = (SawString ? 1 : 0) + (SawNumber ? 1 : 0) + (SawDate ? 1 : 0) + (SawBool ? 1 : 0);
                // A mix of kinds, a formula/error result, nothing sampled at all, or plain text — all
                // four fall back to the string type, the only one able to represent them verbatim.
                if (SawFormulaOrError || kinds != 1 || SawString)
                {
                    return ExcelColumnType.StringColumn;
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
