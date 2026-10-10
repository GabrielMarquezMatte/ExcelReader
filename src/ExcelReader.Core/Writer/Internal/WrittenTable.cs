using ExcelReader.Core.Reader.Internal;

namespace ExcelReader.Core.Writer.Internal
{
    internal sealed record WrittenTable(int Id, string Name, string[] Columns, ExcelTableOptions Options, int HeaderRow, int LastRow)
    {
        internal int FirstColumn
        {
            get
            {
                return Options.FirstColumn;
            }
        }

        internal int LastColumn
        {
            get
            {
                return Options.FirstColumn + Columns.Length - 1;
            }
        }

        internal string Ref
        {
            get
            {
                return TablePart.FormatRange(HeaderRow, FirstColumn, LastRow, LastColumn);
            }
        }
    }
}
