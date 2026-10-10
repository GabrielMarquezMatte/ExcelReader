using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;

namespace ExcelReader.Core.Reader.Internal
{
    internal static class WorkbookTables
    {
        internal static ReadOnlyCollection<ExcelTable> Empty
        {
            get
            {
                return ReadOnlyCollection<ExcelTable>.Empty;
            }
        }

        internal static ReadOnlyCollection<ExcelTable> Build<TSheet>(List<TablePart> parts, ExcelSheetList<TSheet> sheets)
            where TSheet : IExcelSheet
        {
            if (parts.Count == 0)
            {
                return Empty;
            }
            parts.Sort(CompareBySheetThenId);
            ExcelTable[] tables = new ExcelTable[parts.Count];
            for (int i = 0; i < tables.Length; i++)
            {
                tables[i] = new ExcelTable(parts[i], sheets[parts[i].SheetIndex]);
            }
            return Array.AsReadOnly(tables);
        }

        internal static bool TryFind(ReadOnlyCollection<ExcelTable> tables, ReadOnlySpan<char> name, [MaybeNullWhen(false)] out ExcelTable table)
        {
            foreach (ExcelTable candidate in tables)
            {
                if (name.Equals(candidate.Name, StringComparison.OrdinalIgnoreCase))
                {
                    table = candidate;
                    return true;
                }
            }
            table = null;
            return false;
        }

        private static int CompareBySheetThenId(TablePart left, TablePart right)
        {
            int bySheet = left.SheetIndex.CompareTo(right.SheetIndex);
            if (bySheet != 0)
            {
                return bySheet;
            }
            return left.Id.CompareTo(right.Id);
        }
    }
}
