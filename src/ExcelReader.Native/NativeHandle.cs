using ExcelReader.Core.Reader;

namespace ExcelReader.Native
{
    internal sealed class NativeHandle : IDisposable
    {
        internal NativeHandle(IExcelWorkbook workbook)
        {
            Workbook = workbook;
        }

        internal IExcelWorkbook Workbook { get; }

        public void Dispose()
        {
            Workbook.Dispose();
        }
    }
}
