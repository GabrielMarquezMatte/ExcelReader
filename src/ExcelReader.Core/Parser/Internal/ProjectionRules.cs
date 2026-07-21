using System.Globalization;

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

        internal static InvalidOperationException MissingRequiredValue(string name, int row)
        {
            return new InvalidOperationException(
                $"Required column '{name}' has no value in row {row.ToString(CultureInfo.InvariantCulture)}.");
        }
    }
}
