using System.Globalization;

namespace Snippets.Versioning.EnumSwitch
{
    internal static class Snippet
    {
        internal static string Run(Cell cell)
        {
            #region versioning-enum-switch
            string text = cell.Type switch
            {
                CellType.Empty => "",
                CellType.Number when cell.TryGetDouble(out double number) => number.ToString(CultureInfo.InvariantCulture),
                _ => cell.GetString(),
            };
            #endregion
            return text;
        }
    }
}
