using ExcelReader.Core.Reader;

namespace ExcelReader.Native.Reading
{
    internal static unsafe partial class ReadApi
    {
        internal static int ForEachRow(NativeHandle? handle, int sheet,
            delegate* unmanaged<NativeCellView*, int, void*, int> callback, void* userData)
        {
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            NativeApi.ClearLastError();
            int status = ResolveSheet(handle, sheet, out IExcelSheet? resolved);
            if (status != NativeStatus.Ok)
            {
                return status;
            }

            try
            {
                using CellViewBuffer views = new();
                using IExcelRowEnumerator rows = resolved!.GetEnumerator();
                while (rows.MoveNext())
                {
                    Row row = rows.Current;
                    fixed (byte* rowValues = row.RowValues)
                    fixed (byte* shared = row.Shared)
                    fixed (byte* rowBuffer = row.RowBuffer)
                    {
                        NativeCellView* cells = views.Fill(row.Descriptors, rowValues, shared, rowBuffer);
                        status = callback(cells, row.Descriptors.Length, userData);
                    }

                    if (status != NativeStatus.Ok)
                    {
                        return status;
                    }
                }
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                NativeApi.SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }
    }
}
