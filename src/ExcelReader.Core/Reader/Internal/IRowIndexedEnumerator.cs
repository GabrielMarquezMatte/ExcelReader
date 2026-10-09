namespace ExcelReader.Core.Reader.Internal
{
    internal interface IRowIndexedEnumerator
    {
        int RowIndex { get; }

        void EnableRowIndex();
    }
}
