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

        internal static ExcelParseException MissingRequiredValue(string name, int row)
        {
            return new ExcelParseException(row, name);
        }

        internal static ExcelParseException AtRow(ExcelParseException original, long row)
        {
            return original.RawValue.Length == 0
                ? MissingRequiredValue(original.ColumnName, (int)row)
                : new ExcelParseException((int)row, original.ColumnName, original.RawValue);
        }
    }
}
