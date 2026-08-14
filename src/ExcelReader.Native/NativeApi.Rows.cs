using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Native
{
    internal static partial class NativeApi
    {
        internal static int NextRow(NativeHandle? handle, Span<byte> buffer, out int written)
        {
            written = 0;
            if (handle is null)
            {
                return NativeStatus.InvalidHandle;
            }

            ClearLastError();
            try
            {
                if (!handle.HasPending)
                {
                    handle.Rows ??= handle.Reader.GetEnumerator();
                    if (!handle.Rows.MoveNext())
                    {
                        return NativeStatus.Eof;
                    }

                    Row row = handle.Rows.Current;
                    byte[] scratch = handle.Scratch;
                    handle.PendingLength = RowBlob.Serialize(row, ref scratch);
                    handle.Scratch = scratch;
                    handle.HasPending = true;
                }

                written = handle.PendingLength;
                if (buffer.Length < handle.PendingLength)
                {
                    return NativeStatus.BufferTooSmall;
                }

                handle.Scratch.AsSpan(0, handle.PendingLength).CopyTo(buffer);
                handle.HasPending = false;
                return NativeStatus.Ok;
            }
            catch (Exception exception)
            {
                SetLastError(exception.Message);
                return NativeStatus.Error;
            }
        }
    }
}
