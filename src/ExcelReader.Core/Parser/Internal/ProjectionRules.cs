namespace ExcelReader.Core.Parser.Internal
{
    internal static class ProjectionRules
    {
        internal static ProjectionStep ClassifyRow(ref int rowNumber, int headerRow, bool mapBuilt)
        {
            rowNumber++;
            if (rowNumber < headerRow)
            {
                return ProjectionStep.Skip;
            }
            if (rowNumber == headerRow)
            {
                return ProjectionStep.BuildMap;
            }
            return mapBuilt ? ProjectionStep.Yield : ProjectionStep.Stop;
        }

        // A blank required cell is a defect in the file's data, not a caller mistake — the same fault
        // class ExcelParseException already covers for a cell that fails to parse (see CsvEnumerable/
        // SparseRowProjection's other throw sites), so this reuses that type instead of
        // InvalidOperationException, which would tell the caller they made a mistake when the file did.
        internal static ExcelParseException MissingRequiredValue(string name, int row)
        {
            return new ExcelParseException(row, name);
        }
    }
}
